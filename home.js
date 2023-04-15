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

// 注册按钮点击回调函数
enterCodeButton.addEventListener("click", async () => {
  const timeWait = document.getElementById("timeWait").value;
  // 调用Chrome接口取出当前标签页
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  // alert(tab.id);
  // 以当前标签页为上下文，执行setPageBackgroundColor函数
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    args:[{billNoList: gBillNoList, timeWait: timeWait}],
    function: processBillNoList,
  });
});


async function processBillNoList(args) {
  const billNoList = args.billNoList;
  const timeWait = args.timeWait < 100 ? 100:args.timeWait;
  // alert(`processBillNoList list ${billNoList.length} 个元素`);
  for (let i = 0; i < billNoList.length; ++i) {
    const billNo = billNoList[i];
    // openwrt passwall port 调试用
    // var textarea = document.querySelector("#cbid\\.passwall\\.375abe461ba845d3b30f23a7aa506065\\.port");
    var textarea = document.querySelector("#waybillNo");
    if (!textarea) {
      alert("没有查找到编辑框");
      return;
    }
    await new Promise(resolve => setTimeout(resolve, timeWait)); // 等待 500ms
    textarea.value = billNo;
    // var form = document.querySelector("#maincontent > div > form");
    // if (!form) {
    //   alert("没有查找到表单");
    //   return;
    // }
    // form.submit();
    // 如果模拟回车不成功,需要查找form进行提交
    // 创建并触发一个 KeyboardEvent 对象
    var enterKeyEvent = new KeyboardEvent("keydown", {
      key: "Enter",
      bubbles: true,
      cancelable: true,
      shiftKey: false,
    });
    textarea.dispatchEvent(enterKeyEvent);
  }
}
