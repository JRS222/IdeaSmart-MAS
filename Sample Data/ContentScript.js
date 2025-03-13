function wait(b) {
    for (var a = (new Date).getTime(); (new Date).getTime() < a + b; )
        ;
}
function GetDecodedURL() {
    try {
        return decodeURIComponent(window.location.href)
    } catch (b) {
        return console.log("GetDecodedURL Exception: " + b),
        console.log("Returning Non Decoded URL"),
        window.location.href
    }
}
function GetTabTitle() {
    return document.title
}
function SendFileDetailsToAgent(b) {
    var a = {};
    a.URL = GetDecodedURL();
    a.FILE = b;
    chrome.runtime.sendMessage({
        FILE_UPLOAD: a
    });
    console.log("File details sent: ");
    console.log(a);
    wait(500)
}
function SendFileDetail(b) {
    if (null != b && null != b.files && 0 != b.files.length) {
        for (var a = [], c = 0; c < b.files.length; c++) {
            var d = b.files[c].name;
            d += "|";
            d += b.files[c].lastModified.toString();
            d += "|";
            a.push(d)
        }
        SendFileDetailsToAgent(a)
    }
}
async function getFile(b) {
    try {
        return await new Promise( (a, c) => b.file(a, c))
    } catch (a) {
        console.log(a)
    }
}
async function getFolderEntries(b) {
    try {
        return await new Promise( (a, c) => b.readEntries(a, c))
    } catch (a) {
        console.log(a)
    }
}
async function ReadFileNames(b, a, c) {
    console.log("ReadFileNames started");
    var d = await getFolderEntries(b);
    if (d.length) {
        for (var e = 0; e < d.length; ++e)
            if (1 == d[e].isFile) {
                var f = await getFile(d[e])
                  , g = f.name;
                g += "|";
                g += f.lastModifiedDate.valueOf().toString();
                g += "|";
                c.push(g);
                console.log("Added filename: " + g + " in list")
            } else
                1 == d[e].isDirectory && (a.push(d[e]),
                console.log("Added subfolder: " + d[e].name + " in list"));
        ReadFileNames(b, a, c)
    } else
        0 == a.length && 0 < c.length ? SendFileDetailsToAgent(c) : 0 < a.length && (b = a.shift(),
        console.log("Started sub-folder: " + b.name),
        ReadFileNames(b.createReader(), a, c));
    console.log("ReadFileNames completed")
}
function EnumAndSendFolderDetails(b, a) {
    console.log("EnumAndSendFolderDetails started");
    for (var c = 0; c < b.length; c++) {
        var d = b[c].webkitGetAsEntry();
        d.isFile ? a.add(d.name) : d.isDirectory && (d = d.createReader(),
        filenames = [],
        subDirectories = [],
        ReadFileNames(d, subDirectories, filenames))
    }
    console.log("EnumAndSendFolderDetails completed")
}
function EnumAndSendFileDetails(b, a) {
    console.log("EnumAndSendFileDetails started");
    for (var c = [], d = 0; d < b.length; d++) {
        var e = b[d];
        if (a.has(e.name)) {
            var f = e.name;
            f += "|";
            f += e.lastModified.toString();
            f += "|";
            c.push(f)
        }
    }
    0 < c.length && SendFileDetailsToAgent(c);
    console.log("EnumAndSendFileDetails completed")
}
function onDrop(b) {
    try {
        if (console.log("onDrop started"),
        null != b.dataTransfer && null != b.dataTransfer.items && 0 != b.dataTransfer.items.length) {
            var a = new Set;
            EnumAndSendFolderDetails(b.dataTransfer.items, a);
            EnumAndSendFileDetails(b.dataTransfer.files, a);
            console.log("onDrop completed")
        }
    } catch (c) {
        console.log(c.message)
    }
}
function onChange(b) {
    try {
        console.log("onChange started"),
        SendFileDetail(b.target),
        console.log("onChange completed")
    } catch (a) {
        console.log(a.message)
    }
}
function onPaste(b) {
    try {
        console.log("onPaste started");
        var a = b.clipboardData;
        if (null != a && null != a.items && 0 != a.items.length) {
            if (a.items[0].webkitGetAsEntry()) {
                let c = new Set;
                EnumAndSendFolderDetails(a.items, c);
                EnumAndSendFileDetails(a.files, c)
            }
            console.log("onPaste completed")
        }
    } catch (c) {
        console.log(c.message)
    }
}
function SendPrintOperationDetailsToAgent(b) {
    var a = {};
    a.URL = GetDecodedURL();
    a.TAB_TITLE = GetTabTitle();
    a.PRINT_CONTENT = b;
    chrome.runtime.sendMessage({
        PRINT_OPERATION: a
    });
    console.log("Print operation details sent for url : " + a.URL + "and Table title: " + a.TAB_TITLE);
    wait(500)
}
(function() {
    window.onbeforeprint = function() {
        console.log("beforePrint called");
        var b = document.querySelectorAll("body")
          , a = "";
        for (i = 0; i < b.length; i++)
            a += b[i].outerHTML;
        SendPrintOperationDetailsToAgent(a)
    }
}
)();
function addDLPEventListeners() {
    document.addEventListener("drop", onDrop, !0);
    document.addEventListener("change", onChange, !0);
    document.addEventListener("paste", onPaste, !0)
}
addDLPEventListeners();
setTimeout(addDLPEventListeners, 1E3);
