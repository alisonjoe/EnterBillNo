// 通过ID找到按钮
const enterCodeButton = document.getElementById("enterCode");
const inputElement = document.getElementById('docpicker'); // 获取 HTML input 元素
inputElement.addEventListener('change', handleFileSelect);

const gBillNoList = [];


// 添加文件后分析要处理的数据
function handleFileSelect(event) {
  // gBillNoList.length = 0
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = e.target.result;
    const workbook = XLSX.read(data, { type: 'binary' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    var idx = 0;
    // 读取第一列的数据
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let i = range.s.r; i <= range.e.r; i++) {
      const cellAddress = { c: 0, r: i }; // 第一列的列索引为0
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      if (cell == null) {
        alert(`第 ${i+1} 行数据异常,请查证`);
        continue
      }
      gBillNoList.push(cell.v);
      idx++;
    }
    var showMsg = `总共 ${idx} 条数据要处理,请确认: (注意: 只展示前后十条数据,如果数据不足20条,展示会有重复,可以忽略)\r\n`;
    const firstTen = gBillNoList.length > 20 ? gBillNoList.slice(0, 10) : gBillNoList;
    const lastTen = gBillNoList.length > 20 ? gBillNoList.slice(-10) : gBillNoList;

    for (let i = 0; i < firstTen.length; i++) { 
      showMsg += firstTen[i] + `\r\n`;
    }
    showMsg += "...";
    showMsg += "\r\n";
    for (let i = 0; i < lastTen.length; i++) {
      showMsg += lastTen[i] + `\r\n`;
    }
    document.getElementById('textshow').value = showMsg;
    // alert(gBillNoList.length);
  };
  reader.readAsBinaryString(file);
}


enterCodeButton.addEventListener("click", function() {
  var timeWait = document.getElementById("timeWait").value;
  console.log("addEventListener begin");
  chrome.tabs.query({ active: true, currentWindow: true }, function(tabs) {
    var tab = tabs[0];
    chrome.tabs.executeScript(tab.id, {
      code: "(" + fillBillNo.toString() + ")({ billNoList: " + JSON.stringify(gBillNoList) + ", timeWait: " + timeWait + " });"
    }, function() {
      // 回调函数执行完毕后再进行异步操作
      console.log('所有运单已经处理完成！');
    });
  });
});


function fillBillNo(args) {
  console.log('fillBillNo begin！');
  console.log(document.title);
  // const targetNode = document.querySelector("#cbid\\.passwall\\.375abe461ba845d3b30f23a7aa506065\\.port");
  const targetNode = document.querySelector("#waybillNo");
  if (targetNode) {
    console.log('目标节点已经存在！');
    for (var i = 0; i < args.billNoList.length; ++i) {
      var billNo = args.billNoList[i];
      console.log('processBillNoList ' + billNo);

      setTimeout((function (billNo, targetNode) {
        return function () {
          targetNode.value = billNo;

          var enterKeyEvent = new KeyboardEvent("keydown", {
            key: "Enter",
            bubbles: true,
            cancelable: true,
            shiftKey: false
          });

          targetNode.dispatchEvent(enterKeyEvent);
        };
      })(billNo, targetNode), i * args.timeWait);
    }
    console.log('processBillNoList end');
  } else {
    console.log('目标节点不存在！');
    alert('目标节点不存在！等待加载成功或刷新重试');
    const observer = new MutationObserver(function (mutations) {
      mutations.forEach(function (mutation) {
        if (mutation.type === 'childList') {
          const textArea = mutation.target.querySelector('textarea');
          if (textArea) {
            alert('文本编辑框已经加载成功！');
            observer.disconnect();
            for (var i = 0; i < args.billNoList.length; ++i) {
              var billNo = args.billNoList[i];
              console.log('processBillNoList ' + billNo);

              setTimeout((function (billNo, targetNode) {
                return function () {
                  targetNode.value = billNo;

                  var enterKeyEvent = new KeyboardEvent("keydown", {
                    key: "Enter",
                    bubbles: true,
                    cancelable: true,
                    shiftKey: false
                  });

                  targetNode.dispatchEvent(enterKeyEvent);
                };
              })(billNo, targetNode), i * args.timeWait);
            }
          }
        }
      });
    });
    console.log('fillBillNo listen！');

    const config = { childList: true, subtree: true };
    observer.observe(document.body, config);
  }
}
