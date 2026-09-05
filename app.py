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

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
except ImportError:
    DND_FILES = None
    BaseTk = tk.Tk


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
        var outputFolder = new Folder(CFG.outputFolder);
        if (!outputFolder.exists && !outputFolder.create()) throw new Error("无法创建输出文件夹");
        var files=[], totalItems=0;
        if(CFG.mode=="main") {
            var mockups=[];
            if(CFG.mockups && CFG.mockups.length) {
                for(var p=0;p<CFG.mockups.length;p++) {
                    var mainMockup=new File(CFG.mockups[p]);
                    if(mainMockup.exists) mockups.push(mainMockup);
                    else failures.push(decodeURI(mainMockup.name)+"：样机文件不存在");
                }
            } else if(CFG.mockup) {
                var legacyMockup=new File(CFG.mockup);
                if(legacyMockup.exists) mockups.push(legacyMockup);
            }
            if(!mockups.length) throw new Error("没有可用的主图 PSD/PSB 样机");
            for(var m=0;m<CFG.materials.length;m++) {
                var material=new File(CFG.materials[m]);
                if(material.exists) files.push(material);
            }
            if(!files.length) throw new Error("没有可用的主图素材");
            for(var p=0;p<mockups.length;p++) {
                try {
                    var updated=processMain(files,mockups[p],outputFolder);
                    successes+=updated; totalItems+=updated;
                } catch(mainItemError) {
                    failures.push(decodeURI(mockups[p].name)+"："+(mainItemError.message||String(mainItemError)));
                }
            }
        } else {
            var mockup = new File(CFG.mockup);
            if (!mockup.exists) throw new Error("样机文件不存在");
            var inputFolder = new Folder(CFG.inputFolder);
            if (!inputFolder.exists) throw new Error("素材文件夹不存在");
            files = inputFolder.getFiles(function(f) {
                return f instanceof File && /\.(png|jpe?g|tif|tiff|psd)$/i.test(f.name);
            });
            files.sort(function(a,b){ return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
            if (!files.length) throw new Error("素材文件夹中没有支持的图片");
            totalItems=files.length;
            for (var i=0; i<files.length; i++) {
                try { processOne(files[i], mockup, outputFolder, i); successes++; }
                catch(e) { failures.push(decodeURI(files[i].name) + "：" + e.message); closeAllWithoutSaving(); }
            }
        }
        return "OK\n" + successes + "\n" + totalItems + "\n" +
               encodeURIComponent(failures.join("\n")) + "\n" +
               encodeURIComponent(exported.join("\n"));
    } catch(e) {
        return "ERR\n" + encodeURIComponent(e.message || String(e));
    } finally { app.displayDialogs = oldDialogs; }

    function processMain(materials,mockupFile,outputFolder) {
        var documentCount=app.documents.length, doc=null, stage="打开样机";
        try {
            doc=app.open(mockupFile);
            stage="查找智能对象";
            var targets=[]; collectSmartObjects(doc,targets);
            if(!targets.length) throw new Error("样机中没有智能对象图层");
            var seen={}, uniqueTargets=[];
            for(var i=0;i<targets.length;i++) {
                var entry=targets[i], key="$"+entry.key;
                if(seen[key]) continue;
                seen[key]=true;
                uniqueTargets.push(entry);
            }
            for(var i=0;i<uniqueTargets.length;i++) {
                var entry=uniqueTargets[i];
                var artFile=materials[i%materials.length];
                stage="替换第 "+(i+1)+" / "+uniqueTargets.length+" 个智能对象“"+entry.layer.name+"”";
                app.activeDocument=doc; doc.activeLayer=entry.layer;
                replaceContents(artFile,CFG.fit||"cover");
                app.activeDocument=doc;
            }
            stage="存储为 Web 所用格式（旧版）JPG";
            var base=safe(strip(decodeURI(mockupFile.name))+"_主图");
            exportWebJpg(doc,uniqueFile(outputFolder,base,".jpg"));
            return uniqueTargets.length;
        } catch(mainError) {
            throw new Error(stage+"："+(mainError.message||String(mainError)));
        } finally {
            while(app.documents.length>documentCount) {
                try { app.activeDocument.close(SaveOptions.DONOTSAVECHANGES); }
                catch(ignore) { break; }
            }
        }
    }
    function collectSmartObjects(container,out) {
        for(var i=0;i<container.layers.length;i++) {
            var l=container.layers[i];
            if(l.typename=="ArtLayer" && l.kind==LayerKind.SMARTOBJECT)
                out.push({layer:l,key:smartObjectKey(l)});
            else if(l.typename=="LayerSet") collectSmartObjects(l,out);
        }
    }
    function smartObjectKey(layer) {
        try {
            var s=stringIDToTypeID, c=charIDToTypeID;
            var ref=new ActionReference(); ref.putIdentifier(c("Lyr "),layer.id);
            var desc=executeActionGet(ref), names=["smartObjectMore","smartObject"];
            for(var i=0;i<names.length;i++) {
                var key=s(names[i]); if(!desc.hasKey(key)) continue;
                var data=desc.getObjectValue(key), props=["ID","placedID","fileReference","link"];
                for(var j=0;j<props.length;j++) {
                    var prop=s(props[j]);
                    if(data.hasKey(prop)) {
                        try { var value=data.getString(prop); if(value) return value; } catch(ignoreString) {}
                        try { var path=data.getPath(prop); if(path) return path.fsName; } catch(ignorePath) {}
                    }
                }
            }
        } catch(ignoreOuter) {}
        return "layer-"+layer.id;
    }

    function processOne(artFile, mockupFile, outputFolder, fileIndex) {
        var stage="打开样机";
        var doc = app.open(mockupFile);
        try {
            stage="查找图像智能对象";
            var target = findSmartObject(doc, CFG.target);
            if (!target) throw new Error('找不到名为“' + CFG.target + '”的智能对象');
            doc.activeLayer = target;
            stage="替换图像智能对象";
            replaceContents(artFile, CFG.fit);
            app.activeDocument = doc;
            if (CFG.textEnabled) {
                stage="查找文字智能对象";
                var textTarget = findSmartObject(doc, CFG.textTarget || "文字");
                if (textTarget && textTarget != target) {
                    try {
                        doc.activeLayer = textTarget;
                        replaceTextSmartObject(strip(decodeURI(artFile.name)));
                    } catch(textError) {
                        failures.push(decodeURI(artFile.name)+"（文字同步已跳过："+(textError.message||String(textError))+"）");
                    } finally { app.activeDocument = doc; }
                }
            }
            stage="导出成品";
            var defaultBase = strip(decodeURI(artFile.name)) + "_" + strip(decodeURI(mockupFile.name));
            var base = safe(applyOutputNaming(defaultBase,fileIndex));
            if (CFG.jpg) exportJpg(doc, uniqueFile(outputFolder, base, ".jpg"), CFG.quality);
            if (CFG.png) exportPng(doc, uniqueFile(outputFolder, base, ".png"));
            if (CFG.slices) {
                try { exportSlicesJpg(doc, outputFolder, base); }
                catch(sliceError) { failures.push(decodeURI(artFile.name)+"（切片已跳过）："+sliceError.message); }
            }
        } catch(processError) {
            throw new Error(stage+"："+(processError.message||String(processError)));
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
        // Photoshop 2020 can reject Copy Merged for some JPG/PNG documents.
        // Flatten a temporary artwork document, then duplicate its layer directly
        // into the Smart Object document. This avoids the clipboard entirely.
        if (art.layers.length > 1) art.flatten();
        var sourceLayer=art.activeLayer;
        var placed=sourceLayer.duplicate(smart, ElementPlacement.PLACEATBEGINNING);
        art.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument=smart;
        placed.name="ARTWORK_REPLACEMENT";
        for(var i=smart.layers.length-1;i>=0;i--){ if(smart.layers[i]!=placed){ try{smart.layers[i].remove();}catch(ignore){} } }
        fitLayer(placed,cw,ch,fit);
        // Some mockups use PNG/JPG files as Smart Object contents. Photoshop
        // 2020 warns that those formats cannot preserve layers on save.
        // The artwork already covers the Smart Object canvas, so flatten the
        // temporary content before saving to keep the original format valid.
        saveSmartDocumentCompat(smart);
        smart.close(SaveOptions.DONOTSAVECHANGES); app.activeDocument=parent;
    }
    function saveSmartDocumentCompat(doc) {
        var lower=doc.name.toLowerCase();
        var target=null;
        try { target=doc.fullName; } catch(ignore) {}
        if(target && /\.png$/i.test(lower)) {
            if(doc.layers.length>1) doc.mergeVisibleLayers();
            var png=new PNGSaveOptions(); png.interlaced=false;
            doc.saveAs(target,png,false,Extension.LOWERCASE);
        } else if(target && /\.jpe?g$/i.test(lower)) {
            doc.flatten();
            var jpg=new JPEGSaveOptions();
            jpg.quality=12; jpg.embedColorProfile=true;
            jpg.formatOptions=FormatOptions.STANDARDBASELINE; jpg.matte=MatteType.WHITE;
            doc.saveAs(target,jpg,false,Extension.LOWERCASE);
        } else {
            doc.save();
        }
    }
    function replaceTextSmartObject(fileBase) {
        var parent=app.activeDocument;
        var textDoc=null, textStage="打开文字智能对象";
        try {
            executeAction(stringIDToTypeID("placedLayerEditContents"),undefined,DialogModes.NO);
            textDoc=app.activeDocument;
            textStage="查找内部文字图层";
            var textLayer=findFirstTextLayer(textDoc);
            if(!textLayer) throw new Error('“'+(CFG.textTarget||'文字')+'”智能对象内没有文字图层');
            textStage="替换并自动排版文字";
            autoLayoutText(textLayer,fileBase,textDoc);
            textStage="保存文字智能对象";
            textDoc.save();
        } catch(textInnerError) {
            throw new Error(textStage+"："+(textInnerError.message||String(textInnerError)));
        } finally {
            try { if(textDoc) textDoc.close(SaveOptions.DONOTSAVECHANGES); } catch(ignore) {}
            app.activeDocument=parent;
        }
    }
    function findFirstTextLayer(container) {
        for(var i=0;i<container.layers.length;i++) {
            var l=container.layers[i];
            if(l.typename=="ArtLayer" && l.kind==LayerKind.TEXT) return l;
            if(l.typename=="LayerSet") { var nested=findFirstTextLayer(l); if(nested) return nested; }
        }
        return null;
    }
    function autoLayoutText(layer,value,doc) {
        // Mixed-style text can make Photoshop 2020 fail when reading size.
        // Write contents first, then contain-fit the editable text layer inside
        // the Smart Object canvas so long names remain completely visible.
        var item=layer.textItem;
        try { item.contents=value; }
        catch(domTextError) { setTextContentsByAction(value); }
        fitLayer(layer,doc.width.as("px"),doc.height.as("px"),"contain");
    }
    function setTextContentsByAction(value) {
        var c=charIDToTypeID;
        var desc=new ActionDescriptor(), ref=new ActionReference(), textDesc=new ActionDescriptor();
        ref.putEnumerated(c("TxLr"),c("Ordn"),c("Trgt"));
        desc.putReference(c("null"),ref);
        textDesc.putString(c("Txt "),value);
        desc.putObject(c("T   "),c("TxLr"),textDesc);
        executeAction(c("setd"),desc,DialogModes.NO);
    }
    function wrapFileName(value,limit) {
        if(value.length<=limit) return value;
        var out=[],line="";
        for(var i=0;i<value.length;i++) {
            line+=value.charAt(i);
            if(line.length>=limit && i<value.length-1) { out.push(line); line=""; }
        }
        if(line) out.push(line); return out.join("\r");
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
    function exportWebJpg(doc,file) {
        app.activeDocument=doc;
        var o=new ExportOptionsSaveForWeb();
        o.format=SaveDocumentType.JPEG;
        o.quality=100;
        o.optimized=true;
        o.interlaced=false;
        o.includeProfile=true;
        o.transparency=false;
        doc.exportDocument(file,ExportType.SAVEFORWEB,o);
        exported.push(file.fsName+" [Web旧版/JPG/品质100/优化/无杂边]");
    }
    function exportSlicesJpg(doc, outputFolder, artworkBase) {
        app.activeDocument=doc;
        var root=new Folder(outputFolder.fsName+"/"+artworkBase+"_切片");
        if(!root.exists && !root.create()) throw new Error("无法创建切片输出文件夹："+root.fsName);
        var jpgFolder=new Folder(root.fsName+"/JPG");
        if(!jpgFolder.exists && !jpgFolder.create()) throw new Error("无法创建 JPG 切片文件夹");
        saveUserSlices(jpgFolder);
        exported.push(jpgFolder.fsName+" [存储为Web旧版/JPG/品质100/优化/无杂边]");
    }
    function saveUserSlices(folder) {
        var c=charIDToTypeID;
        var desc=new ActionDescriptor();
        var opts=new ActionDescriptor();
        opts.putEnumerated(c("Op  "),c("SWOp"),c("OpSa"));
        opts.putBoolean(c("DIDr"),true);
        opts.putPath(c("In  "),folder);
        opts.putEnumerated(c("Fmt "),c("IRFm"),c("JPEG"));
        opts.putBoolean(c("Intr"),false);
        opts.putInteger(c("Qlty"),100);
        opts.putBoolean(c("Optm"),true);
        opts.putBoolean(c("Mtt "),false);
        opts.putBoolean(c("SHTM"),false);
        opts.putBoolean(c("SImg"),true);
        opts.putEnumerated(c("SWsl"),c("STsl"),c("SLUs"));
        desc.putObject(c("Usng"),stringIDToTypeID("SaveForWeb"),opts);
        executeAction(c("Expr"),desc,DialogModes.NO);
    }
    function uniqueFile(folder,base,ext) {
        var f=new File(folder.fsName+"/"+base+ext),n=2;
        while(f.exists){ f=new File(folder.fsName+"/"+base+"_"+n+ext); n++; }
        return f;
    }
    function applyOutputNaming(original,index) {
        var n=CFG.naming;
        if(!n || !n.enabled) return original;
        var value=original;
        if(n.mode=="custom") value=n.text || original;
        else if(n.mode=="insert") {
            var pos=resolvePosition(value,n.position,n.customPosition);
            value=value.substring(0,pos)+(n.text||"")+value.substring(pos);
        } else if(n.mode=="replace") {
            if(n.find) value=value.split(n.find).join(n.replace||"");
        } else if(n.mode=="delete") {
            if(n.deleteType=="digits") value=value.replace(/[0-9]/g,"");
            else if(n.deleteType=="letters") value=value.replace(/[A-Za-z]/g,"");
            else if(n.deleteType=="spaces") value=value.replace(/\s+/g,"");
            else if(n.deleteType=="specified" && n.deleteText) value=value.split(n.deleteText).join("");
        }
        if(n.caseMode=="upper") value=value.toUpperCase();
        else if(n.caseMode=="lower") value=value.toLowerCase();
        else if(n.caseMode=="title") value=value.replace(/(^|[ _-])([a-z])/g,function(a,b,c){return b+c.toUpperCase();});
        if(n.numbering) {
            var token=numberToken(n,index);
            if(n.onlyNumber) value=token;
            else {
                var p=resolvePosition(value,n.numberPosition,n.numberCustomPosition);
                value=value.substring(0,p)+token+value.substring(p);
            }
        }
        value=value.replace(/^\s+|\s+$/g,"");
        return value || ("output_"+(index+1));
    }
    function resolvePosition(text,position,customPos) {
        if(position=="start") return 0;
        if(position=="custom") return Math.max(0,Math.min(text.length,parseInt(customPos,10)||0));
        return text.length;
    }
    function numberToken(n,index) {
        var value=(parseInt(n.start,10)||1)+index*(parseInt(n.step,10)||1);
        if(n.numberType=="letter") return letters(value);
        if(n.numberType=="time") {
            var d=new Date();
            return d.getFullYear()+pad(d.getMonth()+1,2)+pad(d.getDate(),2)+"_"+pad(value,parseInt(n.width,10)||2);
        }
        if(n.numberType=="random") {
            var chars="ABCDEFGHJKLMNPQRSTUVWXYZ23456789",out="",len=Math.max(1,parseInt(n.width,10)||4);
            for(var i=0;i<len;i++) out+=chars.charAt(Math.floor(Math.random()*chars.length));
            return out;
        }
        return pad(value,Math.max(1,parseInt(n.width,10)||1));
    }
    function pad(value,width) { var s=String(value); while(s.length<width)s="0"+s; return s; }
    function letters(value) {
        var n=Math.max(1,value),s="";
        while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}
        return s;
    }
    function strip(n){return n.replace(/\.[^\.]+$/,'');}
    function safe(n){return n.replace(/[\\\/:*?"<>|]/g,'_');}
    function closeAllWithoutSaving(){while(app.documents.length){try{app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}catch(e){break;}}}
})();'''.replace("__CFG__", cfg)


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("900x820")
        self.minsize(800, 700)
        self.events = queue.Queue()
        self.naming = self._default_naming()
        self.main_mockup_files = []
        self.main_material_files = []
        self._build()
        self.after(150, self._poll)

    def _build(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text=APP_TITLE, font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(root, text="调用本机 Photoshop，保留智能对象的透视、变形、蒙版和光影效果。", foreground="#555").pack(anchor="w", pady=(4, 10))
        self.pages = ttk.Notebook(root)
        self.pages.pack(fill="x")
        sku_page = ttk.Frame(self.pages, padding=12)
        main_page = ttk.Frame(self.pages, padding=12)
        self.pages.add(sku_page, text="SKU")
        self.pages.add(main_page, text="主图")
        self._build_sku_page(sku_page)
        self._build_main_page(main_page)
        self.status = tk.StringVar(value="等待开始")
        ttk.Label(root, textvariable=self.status).pack(anchor="w", pady=(10, 0))
        self.log = tk.Text(root, height=11, wrap="word", state="disabled", font=("Microsoft YaHei UI", 9))
        self.log.pack(fill="both", expand=True, pady=(6, 0))

    def _build_sku_page(self, root):
        self.mockup = self._path_row(root, "样机 PSD / PSB", self._choose_mockup)
        self.inputs = self._path_row(root, "素材文件夹", lambda: self._choose_dir(self.inputs))
        self.output = self._path_row(root, "输出文件夹", lambda: self._choose_dir(self.output))
        opts = ttk.LabelFrame(root, text="SKU 处理设置", padding=12)
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
        text_row = ttk.Frame(opts); text_row.pack(fill="x", pady=(12, 0))
        self.text_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(text_row, text="同步素材文件名到文字智能对象", variable=self.text_enabled).pack(side="left")
        ttk.Label(text_row, text="智能对象名：").pack(side="left", padx=(18, 4))
        self.text_target = tk.StringVar(value="文字")
        ttk.Entry(text_row, textvariable=self.text_target, width=12).pack(side="left")
        fmt = ttk.Frame(opts); fmt.pack(fill="x", pady=(12, 0))
        self.jpg = tk.BooleanVar(value=True); self.png = tk.BooleanVar(value=True)
        ttk.Checkbutton(fmt, text="导出 JPG", variable=self.jpg).pack(side="left")
        ttk.Checkbutton(fmt, text="导出 PNG", variable=self.png).pack(side="left", padx=18)
        ttk.Label(fmt, text="JPG 品质：").pack(side="left")
        self.quality = tk.IntVar(value=10)
        ttk.Spinbox(fmt, from_=1, to=12, textvariable=self.quality, width=5).pack(side="left")
        self.slices = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="PSD 有切片时：存储为Web旧版 JPG（品质100/优化/无杂边）", variable=self.slices).pack(anchor="w", pady=(12, 0))
        naming_row = ttk.Frame(opts); naming_row.pack(fill="x", pady=(12, 0))
        ttk.Button(naming_row, text="输出命名设置…", command=self._open_naming).pack(side="left")
        self.naming_summary = tk.StringVar(value="保持默认名称：素材名_样机名")
        ttk.Label(naming_row, textvariable=self.naming_summary, foreground="#555").pack(side="left", padx=12)
        self.run_btn = ttk.Button(root, text="开始批量替换并导出", command=self._start)
        self.run_btn.pack(fill="x", ipady=8, pady=(4, 0))

    def _build_main_page(self, root):
        mockup_row = ttk.Frame(root); mockup_row.pack(fill="both", pady=5)
        ttk.Label(mockup_row, text="主图 PSD / PSB", width=16).pack(side="left", anchor="n", pady=3)
        mockup_box = ttk.Frame(mockup_row); mockup_box.pack(side="left", fill="both", expand=True, padx=8)
        self.main_mockup_list = tk.Listbox(mockup_box, height=4, selectmode="extended")
        mockup_scroll = ttk.Scrollbar(mockup_box, orient="vertical", command=self.main_mockup_list.yview)
        self.main_mockup_list.configure(yscrollcommand=mockup_scroll.set)
        self.main_mockup_list.pack(side="left", fill="both", expand=True)
        mockup_scroll.pack(side="right", fill="y")
        mockup_buttons = ttk.Frame(mockup_row); mockup_buttons.pack(side="right", anchor="n")
        ttk.Button(mockup_buttons, text="选择样机…", command=self._choose_main_mockups).pack(fill="x")
        ttk.Button(mockup_buttons, text="删除选中", command=self._remove_main_mockups).pack(fill="x", pady=5)
        ttk.Button(mockup_buttons, text="清空", command=self._clear_main_mockups).pack(fill="x")
        if DND_FILES:
            self.main_mockup_list.drop_target_register(DND_FILES)
            self.main_mockup_list.dnd_bind("<<Drop>>", self._drop_main_mockups)
            mockup_tip = "可多选 PSD/PSB，也可把样机文件或文件夹拖入上方列表"
        else:
            mockup_tip = "点击“选择样机”可一次选择多个 PSD/PSB"
        ttk.Label(root, text=mockup_tip, foreground="#16834f").pack(anchor="w", padx=(128, 0))
        material_row = ttk.Frame(root); material_row.pack(fill="both", pady=5)
        ttk.Label(material_row, text="主图素材", width=16).pack(side="left", anchor="n", pady=3)
        material_box = ttk.Frame(material_row); material_box.pack(side="left", fill="both", expand=True, padx=8)
        self.main_material_list = tk.Listbox(material_box, height=6, selectmode="extended")
        material_scroll = ttk.Scrollbar(material_box, orient="vertical", command=self.main_material_list.yview)
        self.main_material_list.configure(yscrollcommand=material_scroll.set)
        self.main_material_list.pack(side="left", fill="both", expand=True)
        material_scroll.pack(side="right", fill="y")
        material_buttons = ttk.Frame(material_row); material_buttons.pack(side="right", anchor="n")
        ttk.Button(material_buttons, text="选择图片…", command=self._choose_main_materials).pack(fill="x")
        ttk.Button(material_buttons, text="删除选中", command=self._remove_main_materials).pack(fill="x", pady=5)
        ttk.Button(material_buttons, text="清空", command=self._clear_main_materials).pack(fill="x")
        if DND_FILES:
            self.main_material_list.drop_target_register(DND_FILES)
            self.main_material_list.dnd_bind("<<Drop>>", self._drop_main_materials)
            drop_text = "可将 PNG/JPG/TIFF/PSD 图片或文件夹拖入上方列表"
        else:
            drop_text = "点击“选择图片”添加素材（当前环境未加载拖放组件）"
        ttk.Label(root, text=drop_text, foreground="#16834f").pack(anchor="w", padx=(128, 0))
        self.main_output = self._path_row(root, "输出文件夹", lambda: self._choose_dir(self.main_output))
        opts = ttk.LabelFrame(root, text="主图处理设置", padding=12)
        opts.pack(fill="x", pady=12)
        fit_row = ttk.Frame(opts); fit_row.pack(fill="x")
        ttk.Label(fit_row, text="智能对象适配方式：").pack(side="left")
        self.main_fit = tk.StringVar(value="cover")
        ttk.Combobox(fit_row, textvariable=self.main_fit, state="readonly", width=14,
                     values=("cover", "contain", "stretch")).pack(side="left", padx=8)
        ttk.Label(opts, text="cover=填充画布　contain=完整显示　stretch=拉伸铺满", foreground="#666").pack(anchor="w", pady=(8, 0))
        ttk.Label(opts, text="替换规则：按 PSD 图层顺序替换全部智能对象；链接同步的重复对象自动跳过；素材不足时循环使用。",
                  foreground="#444", wraplength=790).pack(anchor="w", pady=(12, 0))
        ttk.Label(opts, text="默认导出：存储为 Web 所用格式（旧版）JPG，品质 100，优化，无杂边。",
                  foreground="#075d32").pack(anchor="w", pady=(12, 0))
        self.main_run_btn = ttk.Button(root, text="开始替换全部智能对象并导出主图", command=self._start_main)
        self.main_run_btn.pack(fill="x", ipady=8, pady=(4, 0))

    def _path_row(self, parent, label, command):
        frame = ttk.Frame(parent); frame.pack(fill="x", pady=5)
        ttk.Label(frame, text=label, width=16).pack(side="left")
        var = tk.StringVar(); ttk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(frame, text="选择…", command=command).pack(side="right")
        return var

    def _default_naming(self):
        return {"enabled":False,"mode":"custom","text":"","find":"","replace":"","deleteType":"digits",
                "deleteText":"","position":"end","customPosition":0,"caseMode":"unchanged","numbering":False,
                "onlyNumber":False,"numberPosition":"end","numberCustomPosition":0,"numberType":"number",
                "start":1,"step":1,"width":2}

    def _open_naming(self):
        w=tk.Toplevel(self); w.title("输出命名设置"); w.geometry("620x590"); w.transient(self); w.grab_set()
        enabled=tk.BooleanVar(value=self.naming["enabled"])
        ttk.Checkbutton(w,text="启用输出命名规则",variable=enabled).pack(anchor="w",padx=16,pady=(14,8))
        book=ttk.Notebook(w); book.pack(fill="both",expand=True,padx=14)
        vars_={}
        for key,kind in [("text","str"),("find","str"),("replace","str"),("deleteText","str"),
                         ("position","str"),("customPosition","int"),("deleteType","str"),("caseMode","str")]:
            vars_[key]=(tk.IntVar if kind=="int" else tk.StringVar)(value=self.naming[key])
        custom=ttk.Frame(book,padding=14); insert=ttk.Frame(book,padding=14); replace=ttk.Frame(book,padding=14); delete=ttk.Frame(book,padding=14)
        book.add(custom,text="自定义"); book.add(insert,text="插入"); book.add(replace,text="替换"); book.add(delete,text="删除")
        book.select(("custom","insert","replace","delete").index(self.naming.get("mode","custom")))
        ttk.Label(custom,text="新文件名").grid(row=0,column=0,sticky="w",pady=6)
        ttk.Entry(custom,textvariable=vars_["text"],width=48).grid(row=0,column=1,sticky="ew",pady=6)
        ttk.Label(insert,text="插入内容").grid(row=0,column=0,sticky="w",pady=6)
        ttk.Entry(insert,textvariable=vars_["text"],width=48).grid(row=0,column=1,sticky="ew",pady=6)
        ttk.Label(insert,text="位置").grid(row=1,column=0,sticky="w",pady=6)
        ttk.Combobox(insert,textvariable=vars_["position"],state="readonly",values=("start","end","custom"),width=15).grid(row=1,column=1,sticky="w")
        ttk.Label(insert,text="自定义位置").grid(row=2,column=0,sticky="w",pady=6)
        ttk.Spinbox(insert,from_=0,to=999,textvariable=vars_["customPosition"],width=8).grid(row=2,column=1,sticky="w")
        ttk.Label(replace,text="查找内容").grid(row=0,column=0,sticky="w",pady=6)
        ttk.Entry(replace,textvariable=vars_["find"],width=48).grid(row=0,column=1,pady=6)
        ttk.Label(replace,text="替换内容").grid(row=1,column=0,sticky="w",pady=6)
        ttk.Entry(replace,textvariable=vars_["replace"],width=48).grid(row=1,column=1,pady=6)
        ttk.Label(delete,text="删除样式").grid(row=0,column=0,sticky="w",pady=6)
        ttk.Combobox(delete,textvariable=vars_["deleteType"],state="readonly",values=("digits","letters","spaces","specified"),width=18).grid(row=0,column=1,sticky="w")
        ttk.Label(delete,text="指定文字").grid(row=1,column=0,sticky="w",pady=6)
        ttk.Entry(delete,textvariable=vars_["deleteText"],width=36).grid(row=1,column=1,sticky="w")
        common=ttk.LabelFrame(w,text="通用设置",padding=12); common.pack(fill="x",padx=14,pady=10)
        ttk.Label(common,text="大小写").grid(row=0,column=0,sticky="w")
        ttk.Combobox(common,textvariable=vars_["caseMode"],state="readonly",values=("unchanged","upper","lower","title"),width=15).grid(row=0,column=1,sticky="w",padx=6)
        num=tk.BooleanVar(value=self.naming["numbering"]); only=tk.BooleanVar(value=self.naming["onlyNumber"])
        npos=tk.StringVar(value=self.naming["numberPosition"]); ncustom=tk.IntVar(value=self.naming["numberCustomPosition"])
        ntype=tk.StringVar(value=self.naming["numberType"]); start=tk.IntVar(value=self.naming["start"]); step=tk.IntVar(value=self.naming["step"]); width=tk.IntVar(value=self.naming["width"])
        ttk.Checkbutton(common,text="编号设置",variable=num).grid(row=1,column=0,sticky="w",pady=(10,4))
        ttk.Checkbutton(common,text="仅使用编号作为文件名",variable=only).grid(row=1,column=1,columnspan=3,sticky="w",pady=(10,4))
        ttk.Label(common,text="位置").grid(row=2,column=0,sticky="w"); ttk.Combobox(common,textvariable=npos,state="readonly",values=("start","end","custom"),width=10).grid(row=2,column=1,sticky="w")
        ttk.Label(common,text="自定义位置").grid(row=2,column=2,sticky="e"); ttk.Spinbox(common,from_=0,to=999,textvariable=ncustom,width=6).grid(row=2,column=3,sticky="w")
        ttk.Label(common,text="类型").grid(row=3,column=0,sticky="w",pady=6); ttk.Combobox(common,textvariable=ntype,state="readonly",values=("number","letter","random","time"),width=10).grid(row=3,column=1,sticky="w")
        ttk.Label(common,text="起始").grid(row=3,column=2,sticky="e"); ttk.Spinbox(common,from_=1,to=999999,textvariable=start,width=7).grid(row=3,column=3,sticky="w")
        ttk.Label(common,text="增量").grid(row=4,column=0,sticky="w"); ttk.Spinbox(common,from_=1,to=999,textvariable=step,width=7).grid(row=4,column=1,sticky="w")
        ttk.Label(common,text="位数").grid(row=4,column=2,sticky="e"); ttk.Spinbox(common,from_=1,to=12,textvariable=width,width=7).grid(row=4,column=3,sticky="w")
        buttons=ttk.Frame(w); buttons.pack(fill="x",padx=14,pady=(0,12))
        def save():
            mode=("custom","insert","replace","delete")[book.index(book.select())]
            self.naming={"enabled":enabled.get(),"mode":mode,"text":vars_["text"].get(),"find":vars_["find"].get(),
                         "replace":vars_["replace"].get(),"deleteType":vars_["deleteType"].get(),"deleteText":vars_["deleteText"].get(),
                         "position":vars_["position"].get(),"customPosition":vars_["customPosition"].get(),"caseMode":vars_["caseMode"].get(),
                         "numbering":num.get(),"onlyNumber":only.get(),"numberPosition":npos.get(),"numberCustomPosition":ncustom.get(),
                         "numberType":ntype.get(),"start":start.get(),"step":step.get(),"width":width.get()}
            self.naming_summary.set(("已启用："+{"custom":"自定义","insert":"插入","replace":"替换","delete":"删除"}[mode]) if enabled.get() else "保持默认名称：素材名_样机名")
            w.destroy()
        ttk.Button(buttons,text="取消",command=w.destroy).pack(side="right")
        ttk.Button(buttons,text="保存设置",command=save).pack(side="right",padx=8)

    def _choose_mockup(self):
        p = filedialog.askopenfilename(title="选择样机", filetypes=[("Photoshop 样机", "*.psd *.psb"), ("所有文件", "*.*")])
        if p: self.mockup.set(p)

    def _choose_main_mockups(self):
        paths = filedialog.askopenfilenames(
            title="选择主图样机（可多选）",
            filetypes=[("Photoshop 样机", "*.psd *.psb"), ("所有文件", "*.*")],
        )
        self._add_main_mockups(paths)

    def _drop_main_mockups(self, event):
        try:
            paths = self.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        self._add_main_mockups(paths)
        return "break"

    def _add_main_mockups(self, paths):
        supported = {".psd", ".psb"}
        additions = []
        for raw in paths:
            path = Path(str(raw).strip().strip("{}"))
            if path.is_dir():
                additions.extend(p for p in sorted(path.iterdir(), key=lambda x: x.name.lower())
                                 if p.is_file() and p.suffix.lower() in supported)
            elif path.is_file() and path.suffix.lower() in supported:
                additions.append(path)
        known = {os.path.normcase(str(Path(p).resolve())) for p in self.main_mockup_files}
        for path in additions:
            resolved = str(path.resolve())
            key = os.path.normcase(resolved)
            if key not in known:
                self.main_mockup_files.append(resolved)
                known.add(key)
        self._refresh_main_mockups()

    def _refresh_main_mockups(self):
        self.main_mockup_list.delete(0, "end")
        for index, path in enumerate(self.main_mockup_files, 1):
            self.main_mockup_list.insert("end", f"{index:02d}. {Path(path).name}")

    def _remove_main_mockups(self):
        selected = set(self.main_mockup_list.curselection())
        self.main_mockup_files = [p for i, p in enumerate(self.main_mockup_files) if i not in selected]
        self._refresh_main_mockups()

    def _clear_main_mockups(self):
        self.main_mockup_files.clear()
        self._refresh_main_mockups()

    def _choose_main_materials(self):
        paths = filedialog.askopenfilenames(
            title="选择主图素材（可多选）",
            filetypes=[("支持的图片", "*.png *.jpg *.jpeg *.tif *.tiff *.psd"), ("所有文件", "*.*")],
        )
        self._add_main_materials(paths)

    def _drop_main_materials(self, event):
        try:
            paths = self.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        self._add_main_materials(paths)
        return "break"

    def _add_main_materials(self, paths):
        supported = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".psd"}
        additions = []
        for raw in paths:
            path = Path(str(raw).strip().strip("{}"))
            if path.is_dir():
                additions.extend(p for p in sorted(path.iterdir(), key=lambda x: x.name.lower())
                                 if p.is_file() and p.suffix.lower() in supported)
            elif path.is_file() and path.suffix.lower() in supported:
                additions.append(path)
        known = {os.path.normcase(str(Path(p).resolve())) for p in self.main_material_files}
        for path in additions:
            resolved = str(path.resolve())
            key = os.path.normcase(resolved)
            if key not in known:
                self.main_material_files.append(resolved)
                known.add(key)
        self._refresh_main_materials()

    def _refresh_main_materials(self):
        self.main_material_list.delete(0, "end")
        for index, path in enumerate(self.main_material_files, 1):
            self.main_material_list.insert("end", f"{index:02d}. {Path(path).name}")

    def _remove_main_materials(self):
        selected = set(self.main_material_list.curselection())
        self.main_material_files = [p for i, p in enumerate(self.main_material_files) if i not in selected]
        self._refresh_main_materials()

    def _clear_main_materials(self):
        self.main_material_files.clear()
        self._refresh_main_materials()

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
        if self.text_enabled.get() and not self.text_target.get().strip(): return "文字智能对象名不能为空"
        if not (self.jpg.get() or self.png.get()): return "请至少选择一种导出格式"
        return None

    def _start(self):
        err = self._validate()
        if err: messagebox.showwarning(APP_TITLE, err); return
        Path(self.output.get()).mkdir(parents=True, exist_ok=True)
        settings = {"mode":"sku", "mockup":str(Path(self.mockup.get()).resolve()), "inputFolder":str(Path(self.inputs.get()).resolve()),
                    "outputFolder":str(Path(self.output.get()).resolve()), "target":self.target.get().strip(),
                    "textEnabled":self.text_enabled.get(), "textTarget":self.text_target.get().strip(),
                    "fit":self.fit.get(), "jpg":self.jpg.get(), "png":self.png.get(), "slices":self.slices.get(),
                    "quality":max(1,min(12,self.quality.get())), "naming":self.naming}
        self._set_running(True); self.status.set("正在连接 Photoshop…")
        self._append("开始 SKU 任务：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        threading.Thread(target=self._worker, args=(settings,), daemon=True).start()

    def _validate_main(self):
        if not self.main_mockup_files: return "请至少选择一个主图 PSD/PSB 样机"
        if not self.main_material_files: return "请选择或拖入至少一张主图素材"
        if not self.main_output.get(): return "请选择主图输出文件夹"
        missing_mockups = [p for p in self.main_mockup_files if not Path(p).is_file()]
        if missing_mockups: return "有主图样机已被移动或删除，请从列表删除后重新添加"
        missing = [p for p in self.main_material_files if not Path(p).is_file()]
        if missing: return "有素材文件已被移动或删除，请清空后重新添加"
        return None

    def _start_main(self):
        err = self._validate_main()
        if err: messagebox.showwarning(APP_TITLE, err); return
        Path(self.main_output.get()).mkdir(parents=True, exist_ok=True)
        settings = {
            "mode":"main", "mockups":[str(Path(p).resolve()) for p in self.main_mockup_files],
            "materials":[str(Path(p).resolve()) for p in self.main_material_files],
            "outputFolder":str(Path(self.main_output.get()).resolve()),
            "fit":self.main_fit.get(), "naming":{"enabled":False},
        }
        self._set_running(True); self.status.set("正在连接 Photoshop…")
        self._append(f"开始主图任务：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}；样机 {len(self.main_mockup_files)} 个，素材 {len(self.main_material_files)} 个")
        threading.Thread(target=self._worker, args=(settings,), daemon=True).start()

    def _set_running(self, running):
        state = "disabled" if running else "normal"
        self.run_btn.configure(state=state)
        self.main_run_btn.configure(state=state)

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
        name = "PSD主图智能对象替换_日志.txt" if settings.get("mode") == "main" else "PSD样机批量替换_日志.txt"
        path = Path(settings["outputFolder"]) / name
        with path.open("a", encoding="utf-8") as f:
            f.write("\n" + "="*60 + "\n" + "\n".join(lines) + "\n")

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "status": self.status.set(event[1])
                elif event[0] == "done":
                    r, s = event[1], event[2]; self._set_running(False)
                    if r.get("ok"):
                        if s.get("mode") == "main":
                            total_mockups=len(s.get("mockups", [])) or 1
                            completed_mockups=len(r.get("exported", []))
                            lines=[f"主图完成：样机成功 {completed_mockups} / 共 {total_mockups}；已更新智能对象 {r['success']} 个"]
                        else:
                            lines=[f"SKU 完成：成功 {r['success']} / 共 {r['total']}"]
                        lines += ["失败："+x for x in r.get("failures",[])]
                        lines += ["导出："+x for x in r.get("exported",[])]
                        self.status.set(lines[0]); [self._append(x) for x in lines]; self._write_log(s, lines)
                        messagebox.showinfo(APP_TITLE, lines[0] + (f"\n失败 {len(r.get('failures',[]))} 个，详情见日志。" if r.get("failures") else ""))
                    else:
                        msg="处理失败："+r.get("error","未知错误"); self.status.set(msg); self._append(msg); self._write_log(s,[msg]); messagebox.showerror(APP_TITLE,msg)
                elif event[0] == "error":
                    msg="无法运行："+event[1]; self._set_running(False); self.status.set(msg); self._append(msg)
                    self._write_log(event[3],[msg,event[2]])
                    messagebox.showerror(APP_TITLE, msg + "\n\n请确认 Windows 已安装 Photoshop，并先正常启动一次。")
        except queue.Empty: pass
        self.after(150, self._poll)


def self_test():
    sample = {"mockup": "C:/测试/样机.psd", "inputFolder": "C:/测试/素材",
              "outputFolder": "C:/测试/输出", "target": "样机", "fit": "cover",
              "textEnabled": True, "textTarget": "文字",
              "jpg": True, "png": True, "slices": True, "quality": 10,
              "naming": {"enabled":True,"mode":"custom","text":"酒红","caseMode":"unchanged","numbering":True,
                         "onlyNumber":False,"numberPosition":"end","numberType":"number","start":1,"step":1,"width":2}}
    script = build_js(sample)
    checks = ["placedLayerEditContents", "findSmartObject", "exportJpg", "exportPng", "样机",
              "replaceTextSmartObject", "findFirstTextLayer", "autoLayoutText", "wrapFileName", "LayerKind.TEXT"]
    missing = [item for item in checks if item not in script]
    if missing:
        raise RuntimeError("自检失败，缺少：" + ", ".join(missing))
    if "JSON.parse" in script or "JSON.stringify" in script:
        raise RuntimeError("自检失败：脚本仍依赖 Photoshop 2020 不支持的 JSON 对象")
    if ".copy(true)" in script or "smart.paste" in script:
        raise RuntimeError("自检失败：脚本仍使用 Photoshop 2020 可能拒绝的合并拷贝")
    if ".duplicate(smart" not in script:
        raise RuntimeError("自检失败：缺少无剪贴板图层复制流程")
    naming_checks = ["applyOutputNaming", "numberToken", "resolvePosition", "deleteType", "onlyNumber"]
    missing_naming = [item for item in naming_checks if item not in script]
    if missing_naming:
        raise RuntimeError("自检失败，输出命名缺少：" + ", ".join(missing_naming))
    slice_checks = ["exportSlicesJpg", "saveUserSlices", 'c("SLUs")', 'c("JPEG")', 'c("Qlty"),100',
                    'c("Optm"),true', 'c("Mtt "),false', "mergeVisibleLayers", "doc.flatten()",
                    "saveSmartDocumentCompat", "new PNGSaveOptions", "new JPEGSaveOptions"]
    missing_slices = [item for item in slice_checks if item not in script]
    if missing_slices:
        raise RuntimeError("自检失败，切片导出缺少：" + ", ".join(missing_slices))
    main_sample = {"mode":"main", "mockups":["C:/测试/主图1.psd", "C:/测试/主图2.psb"],
                   "materials":["C:/测试/01.png", "C:/测试/02.jpg"],
                   "outputFolder":"C:/测试/主图输出", "fit":"cover", "naming":{"enabled":False}}
    main_script = build_js(main_sample)
    main_checks = ["processMain", "collectSmartObjects", "smartObjectKey", "CFG.mockups", "mockups.length",
                   "uniqueTargets", "materials[i%materials.length]",
                   "exportWebJpg", "ExportOptionsSaveForWeb", "SaveDocumentType.JPEG", "o.quality=100",
                   "o.optimized=true", "ExportType.SAVEFORWEB"]
    missing_main = [item for item in main_checks if item not in main_script]
    if missing_main:
        raise RuntimeError("自检失败，主图流程缺少：" + ", ".join(missing_main))
    if 'fitLayer(layer,doc.width.as("px"),doc.height.as("px"),"contain")' not in script:
        raise RuntimeError("自检失败：SKU 文字未使用 contain 完整显示")
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
