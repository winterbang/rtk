(function(window, document) {
  let X = XLSX;
  function checkFile (file, callback) {
    // let files = e.target.files;
    // let errorDom = document.getElementById("error-table")
    // let perfactDom = document.getElementById("perfact")
    // errorDom.style.display = "none";
    // perfactDom.style.display = "none";
    // let f = files[0];
    function fixdata(data) {
      let o = "", l = 0, w = 10240;
      for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
      o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
      return o;
    }

    let reader = new FileReader();
    let name = file.name;
    console.log(name)
    reader.readAsArrayBuffer(file);
    reader.onload = function(e) {
      var data = e.target.result;
      var arr = fixdata(data);
      var wb = X.read(btoa(arr), {type: 'base64'});
      const sheetNames = wb.SheetNames; // 返回 ['sheet1', 'sheet2']
      // 根据表名获取对应某张表
      const worksheet = wb.Sheets[sheetNames[0]];
      console.log(worksheet);
      if(typeof callback == 'function') callback(worksheet)
    }
  }

  window.checkFile = checkFile
}(window, document))
