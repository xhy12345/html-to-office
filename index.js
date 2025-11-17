import * as XLSX from 'xlsx'
import { asBlob } from "html-docx-js-typescript";
import { saveAs } from "file-saver";
import html2Canvas from "html2canvas";
import JsPDF from "jspdf";

/**
 * table标签转excel表格（xlsx）
 * @param {string} element 
 * @param {string} name 
 */
const exportToExcel = (element, name = '') => {
      const xlsxParam = { raw: true }
      let wb;
      if (typeof element === 'string') {
        wb = XLSX.utils.table_to_book(document.getElementById(element), xlsxParam)
      } else {
        wb = XLSX.utils.book_new();
        element.forEach((item) => XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById(item.eleById), xlsxParam), item.title))
      }
      const fileName = `${name || new Date().getTime()}.xlsx`
      const wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        bookSST: true,
        type: 'array'
      })
      try {
        saveAs(new Blob([wbout], {
          type: 'application/octet-stream'
        }), fileName)
      } catch (e) {
        console.log(e, wbout)
      }
      return wbout
}

/**
 * html转word文档（docx）
 * @param {string} element 
 * @param {string} name 
 */
const htmlToDocx = (element, name = '') => {
    let markdownInfo = document.getElementById(`${element}`).innerHTML;
    asBlob(markdownInfo).then((data) => {
        saveAs(data, `${name || new Date().getTime()}.docx`);
    });
}

/**
 * html转pdf文档（支持多页）
 * @param {string} element 
 * @param {string} name 
 */
const htmlToPdf = (element, name = '') => {
    const ele = document.getElementById(`${element}`);
    let mathInfo = ele.getElementsByTagName('mjx-assistive-mml');
    for(let i = 0; i < mathInfo.length; i++) {
        mathInfo[i].style.opacity = '0';
    }
    html2Canvas(ele, {
      svgRendering: true,
      useCORS: true,
      scale: 2, // 使用设备的像素比
    }).then(canvas => {
      let position = 0 //页面偏移
      const A4W = 595.28 // A4纸宽度
      const A4H = 841.89 // A4纸宽
  
      // 一页PDF可显示的canvas高度
      const pageH = (canvas.width * A4H) / A4W
      // 未分配到PDF的canvas高度
      let unallottedHeight = canvas.height
  
      // canvas对应的PDF宽高
      const imgWidth = A4W - 40
      const imgHeight = (A4W * canvas.height) / canvas.width
  
      const pageData = canvas.toDataURL('image/jpeg', 1.0)
      const pdf = new JsPDF('', 'pt', [A4W, A4H])
  
      // 当canvas高度 未超过 一页PDF可显示的canvas高度，无需分页
      if (unallottedHeight <= pageH) {
        pdf.addImage(pageData, 'JPEG', 20, 0, imgWidth, imgHeight)
        pdf.save(`${name || new Date().getTime()}.pdf`)
        return
      }
  
      while (unallottedHeight > 0) {
        pdf.addImage(pageData, 'JPEG', 20, position, imgWidth, imgHeight)
        unallottedHeight -= pageH
        position -= A4H
        if (unallottedHeight > 0) {
          pdf.addPage()
        }
      }
      pdf.save(`${name || new Date().getTime()}.pdf`)
    })
}

  
export {
    exportToExcel,
    htmlToDocx,
    htmlToPdf
}
