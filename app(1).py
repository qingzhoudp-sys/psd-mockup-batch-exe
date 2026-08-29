import json
import os
import queue
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "PSD 样机智能对象批量替换"
TARGET_LAYER = "样机"


def build_js(settings):
    cfg = json.dumps(settings, ensure_ascii=True, separators=(",", ":"))
    return r'''#target photoshop
(function () {
    // Photoshop 2020 ExtendScript may not expose the global JSON object.
    // A JSON object literal is also valid JavaScript, so embed it directly.
    var CFG = __CFG__;
    var oldDialogs = app.displayDialogs;
    app.displayDialogs = DialogModes.NO;
    var successes = 0, failures = [], exported = [];
    try {
        var mockup = new File(CFG.mockup);
        var inputFolder = new Folder(CFG.inputFolder);
        var outputFolder = new Folder(CFG.outputFolder);
        if (!mockup.exists) throw new Error("样机文件不存在");
        if (!inputFolder.exists) throw new Error("素材文件夹不存在");
        if (!outputFolder.exists && !outputFolder.create()) throw new Error("无法创建输出文件夹");
        var files = inputFolder.getFiles(function(f) {
            return f instanceof File && /\.(png|jpe?g|tif|tiff|psd)$/i.test(f.name);
        });
        files.sort(function(a,b){ return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
        if (!files.length) throw new Error("素材文件夹中没有支持的图片");
        for (var i=0; i<files.length; i++) {
            try { processOne(files[i], mockup, outputFolder); successes++; }
            catch(e) { failures.push(decodeURI(files[i].name) + "：" + e.message); closeAllWithoutSaving(); }
        }
        return "OK\n" + successes + "\n" + files.length + "\n" +
               encodeURIComponent(failures.join("\n")) + "\n" +
               encodeURIComponent(exported.join("\n"));
    } catch(e) {
        return "ERR\n" + encodeURIComponent(e.message || String(e));
    } finally { app.displayDialogs = oldDialogs; }

    function processOne(artFile, mockupFile, outputFolder) {
        var doc = app.open(mockupFile);
        try {
            var target = findSmartObject(doc, CFG.target);
            if (!target) throw new Error('找不到名为“' + CFG.target + '”的智能对象');
            doc.activeLayer = target;
            replaceContents(artFile, CFG.fit);
            app.activeDocument = doc;
            var base = safe(strip(decodeURI(artFile.name)) + "_" + strip(decodeURI(mockupFile.name)));
            if (CFG.jpg) exportJpg(doc, uniqueFile(outputFolder, base, ".jpg"), CFG.quality);
            if (CFG.png) exportPng(doc, uniqueFile(outputFolder, base, ".png"));
        } finally {
            try { app.activeDocument = doc; doc.close(SaveOptions.DONOTSAVECHANGES); } catch(ignore) {}
        }
    }
    function findSmartObject(container, name) {
        for (var i=0; i<container.layers.length; i++) {
            var l=container.layers[i];
            if (l.typename=="ArtLayer" && l.kind==LayerKind.SMARTOBJECT && l.name==name) return l;
            if (l.typename=="LayerSet") { var nested=findSmartObject(l,name); if(nested) return nested; }
        }
        return null;
    }
    function replaceContents(artFile, fit) {
        var parent=app.activeDocument;
        executeAction(stringIDToTypeID("placedLayerEditContents"), undefined, DialogModes.NO);
        var smart=app.activeDocument, cw=smart.width.as("px"), ch=smart.height.as("px");
        var art=app.open(artFile);
        art.selection.selectAll(); art.selection.copy(true); art.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument=smart;
        var placed=smart.paste(); placed.name="ARTWORK_REPLACEMENT";
        for(var i=smart.layers.length-1;i>=0;i--){ if(smart.layers[i]!=placed){ try{smart.layers[i].remove();}catch(ignore){} } }
        fitLayer(placed,cw,ch,fit);
        smart.save(); smart.close(SaveOptions.SAVECHANGES); app.activeDocument=parent;
    }
    function fitLayer(layer,cw,ch,mode) {
        var b=layer.bounds,w=b[2].as("px")-b[0].as("px"),h=b[3].as("px")-b[1].as("px");
        if(w<=0||h<=0) throw new Error("素材没有有效像素");
        if(mode=="stretch") layer.resize(cw/w*100,ch/h*100,AnchorPosition.MIDDLECENTER);
        else { var s=mode=="contain"?Math.min(cw/w,ch/h):Math.max(cw/w,ch/h); layer.resize(s*100,s*100,AnchorPosition.MIDDLECENTER); }
        b=layer.bounds; var cx=(b[0].as("px")+b[2].as("px"))/2,cy=(b[1].as("px")+b[3].as("px"))/2;
        layer.translate(cw/2-cx,ch/2-cy);
    }
    function exportJpg(doc,file,q) {
        var o=new JPEGSaveOptions(); o.quality=q; o.embedColorProfile=true; o.formatOptions=FormatOptions.STANDARDBASELINE; o.matte=MatteType.WHITE;
        doc.saveAs(file,o,true,Extension.LOWERCASE); exported.push(file.fsName);
    }
    function exportPng(doc,file) {
        var o=new PNGSaveOptions(); o.interlaced=false; doc.saveAs(file,o,true,Extension.LOWERCASE); exported.push(file.fsName);
    }
    function uniqueFile(folder,base,ext) {
        var f=new File(folder.fsName+"/"+base+ext),n=2;
        while(f.exists){ f=new File(folder.fsName+"/"+base+"_"+n+ext); n++; }
        return f;
    }
    function strip(n){return n.replace(/\.[^\.]+$/,'');}
    function safe(n){return n.replace(/[\\\/:*?"<>|]/g,'_');}
    function closeAllWithoutSaving(){while(app.documents.length){try{app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}catch(e){break;}}}
})();'''.replace("__CFG__", cfg)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x600")
        self.minsize(700, 540)
        self.events = queue.Queue()
        self._build()
        self.after(150, self._poll)

    def _build(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text=APP_TITLE, font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(root, text="调用本机 Photoshop，保留智能对象的透视、变形、蒙版和光影效果。", foreground="#555").pack(anchor="w", pady=(4, 16))
        self.mockup = self._path_row(root, "样机 PSD / PSB", self._choose_mockup)
        self.inputs = self._path_row(root, "素材文件夹", lambda: self._choose_dir(self.inputs))
        self.output = self._path_row(root, "输出文件夹", lambda: self._choose_dir(self.output))
        opts = ttk.LabelFrame(root, text="处理设置", padding=12)
        opts.pack(fill="x", pady=12)
        row = ttk.Frame(opts); row.pack(fill="x")
        ttk.Label(row, text="智能对象图层名：").pack(side="left")
        self.target = tk.StringVar(value=TARGET_LAYER)
        ttk.Entry(row, textvariable=self.target, width=18).pack(side="left", padx=(4, 24))
        ttk.Label(row, text="适配方式：").pack(side="left")
        self.fit = tk.StringVar(value="cover")
        ttk.Combobox(row, textvariable=self.fit, state="readonly", width=14,
                     values=("cover", "contain", "stretch")).pack(side="left")
        ttk.Label(opts, text="cover=填充画布　contain=完整显示　stretch=拉伸铺满", foreground="#666").pack(anchor="w", pady=(8, 0))
        fmt = ttk.Frame(opts); fmt.pack(fill="x", pady=(12, 0))
        self.jpg = tk.BooleanVar(value=True); self.png = tk.BooleanVar(value=True)
        ttk.Checkbutton(fmt, text="导出 JPG", variable=self.jpg).pack(side="left")
        ttk.Checkbutton(fmt, text="导出 PNG", variable=self.png).pack(side="left", padx=18)
        ttk.Label(fmt, text="JPG 品质：").pack(side="left")
        self.quality = tk.IntVar(value=10)
        ttk.Spinbox(fmt, from_=1, to=12, textvariable=self.quality, width=5).pack(side="left")
        self.run_btn = ttk.Button(root, text="开始批量替换并导出", command=self._start)
        self.run_btn.pack(fill="x", ipady=8, pady=(4, 10))
        self.status = tk.StringVar(value="等待开始")
        ttk.Label(root, textvariable=self.status).pack(anchor="w")
        self.log = tk.Text(root, height=11, wrap="word", state="disabled", font=("Microsoft YaHei UI", 9))
        self.log.pack(fill="both", expand=True, pady=(6, 0))

    def _path_row(self, parent, label, command):
        frame = ttk.Frame(parent); frame.pack(fill="x", pady=5)
        ttk.Label(frame, text=label, width=16).pack(side="left")
        var = tk.StringVar(); ttk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(frame, text="选择…", command=command).pack(side="right")
        return var

    def _choose_mockup(self):
        p = filedialog.askopenfilename(title="选择样机", filetypes=[("Photoshop 样机", "*.psd *.psb"), ("所有文件", "*.*")])
        if p: self.mockup.set(p)

    def _choose_dir(self, var):
        p = filedialog.askdirectory()
        if p: var.set(p)

    def _append(self, text):
        self.log.configure(state="normal"); self.log.insert("end", text + "\n"); self.log.see("end"); self.log.configure(state="disabled")

    def _validate(self):
        if not Path(self.mockup.get()).is_file(): return "请选择有效的 PSD/PSB 样机文件"
        if not Path(self.inputs.get()).is_dir(): return "请选择有效的素材文件夹"
        if not self.output.get(): return "请选择输出文件夹"
        if not self.target.get().strip(): return "智能对象图层名不能为空"
        if not (self.jpg.get() or self.png.get()): return "请至少选择一种导出格式"
        return None

    def _start(self):
        err = self._validate()
        if err: messagebox.showwarning(APP_TITLE, err); return
        Path(self.output.get()).mkdir(parents=True, exist_ok=True)
        settings = {"mockup":str(Path(self.mockup.get()).resolve()), "inputFolder":str(Path(self.inputs.get()).resolve()),
                    "outputFolder":str(Path(self.output.get()).resolve()), "target":self.target.get().strip(),
                    "fit":self.fit.get(), "jpg":self.jpg.get(), "png":self.png.get(), "quality":max(1,min(12,self.quality.get()))}
        self.run_btn.configure(state="disabled"); self.status.set("正在连接 Photoshop…")
        self._append("开始任务：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        threading.Thread(target=self._worker, args=(settings,), daemon=True).start()

    def _worker(self, settings):
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            try:
                ps = win32com.client.Dispatch("Photoshop.Application")
                ps.Visible = True
                self.events.put(("status", "Photoshop 正在批量处理，请勿关闭…"))
                raw = ps.DoJavaScript(build_js(settings))
                result = parse_photoshop_result(raw)
            finally: pythoncom.CoUninitialize()
            self.events.put(("done", result, settings))
        except Exception as exc:
            self.events.put(("error", str(exc), traceback.format_exc(), settings))

    def _write_log(self, settings, lines):
        path = Path(settings["outputFolder"]) / "PSD样机批量替换_日志.txt"
        with path.open("a", encoding="utf-8") as f:
            f.write("\n" + "="*60 + "\n" + "\n".join(lines) + "\n")

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "status": self.status.set(event[1])
                elif event[0] == "done":
                    r, s = event[1], event[2]; self.run_btn.configure(state="normal")
                    if r.get("ok"):
                        lines=[f"完成：成功 {r['success']} / 共 {r['total']}"]
                        lines += ["失败："+x for x in r.get("failures",[])]
                        lines += ["导出："+x for x in r.get("exported",[])]
                        self.status.set(lines[0]); [self._append(x) for x in lines]; self._write_log(s, lines)
                        messagebox.showinfo(APP_TITLE, lines[0] + (f"\n失败 {len(r.get('failures',[]))} 个，详情见日志。" if r.get("failures") else ""))
                    else:
                        msg="处理失败："+r.get("error","未知错误"); self.status.set(msg); self._append(msg); self._write_log(s,[msg]); messagebox.showerror(APP_TITLE,msg)
                elif event[0] == "error":
                    msg="无法运行："+event[1]; self.run_btn.configure(state="normal"); self.status.set(msg); self._append(msg)
                    self._write_log(event[3],[msg,event[2]])
                    messagebox.showerror(APP_TITLE, msg + "\n\n请确认 Windows 已安装 Photoshop，并先正常启动一次。")
        except queue.Empty: pass
        self.after(150, self._poll)


def self_test():
    sample = {"mockup": "C:/测试/样机.psd", "inputFolder": "C:/测试/素材",
              "outputFolder": "C:/测试/输出", "target": "样机", "fit": "cover",
              "jpg": True, "png": True, "quality": 10}
    script = build_js(sample)
    checks = ["placedLayerEditContents", "findSmartObject", "exportJpg", "exportPng", "样机"]
    missing = [item for item in checks if item not in script]
    if missing:
        raise RuntimeError("自检失败，缺少：" + ", ".join(missing))
    if "JSON.parse" in script or "JSON.stringify" in script:
        raise RuntimeError("自检失败：脚本仍依赖 Photoshop 2020 不支持的 JSON 对象")
    ok = parse_photoshop_result("OK\n2\n2\n\nC%3A%2Fout%2Fa.jpg%0AC%3A%2Fout%2Fb.png")
    if not ok.get("ok") or ok.get("success") != 2 or len(ok.get("exported", [])) != 2:
        raise RuntimeError("自检失败：Photoshop 2020 文本结果解析异常")
    err = parse_photoshop_result("ERR\n%E6%B5%8B%E8%AF%95%E9%94%99%E8%AF%AF")
    if err.get("ok") or not err.get("error"):
        raise RuntimeError("自检失败：错误结果解析异常")
    print("SELF_TEST_OK")


def parse_photoshop_result(raw):
    if raw is None:
        raise RuntimeError("Photoshop 没有返回处理结果")
    text = str(raw).replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    if lines and lines[0] == "OK" and len(lines) >= 5:
        failures_text = unquote(lines[3])
        exported_text = unquote("\n".join(lines[4:]))
        return {
            "ok": True,
            "success": int(lines[1]),
            "total": int(lines[2]),
            "failures": [x for x in failures_text.split("\n") if x],
            "exported": [x for x in exported_text.split("\n") if x],
        }
    if lines and lines[0] == "ERR":
        return {"ok": False, "error": unquote("\n".join(lines[1:]))}
    raise RuntimeError("无法识别 Photoshop 返回结果：" + text[:500])


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        App().mainloop()
