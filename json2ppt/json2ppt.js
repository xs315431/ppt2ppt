import pptxgen from "pptxgenjs"; 
import { loadJson } from "../tool/loadJson.js";
// 1. Create a new Presentation

let pres = new pptxgen();
// 加载完模板文件 需要与PPT大纲内容进行替换
let jsonData = loadJson("../pptjson/output.json") 

let { slides } = jsonData


slides.forEach((item, slideIndex) => {
    let slide = pres.addSlide()
    slide.background = { color: item.backgroundColor, path: item.backgroundImage }
    // console.log(item)
    item.elements.forEach((element, elIndex) => {
        const { type } = element;
        switch (type) {
            case "text":
                if (element.text && element.options) {
                    slide.addText(element.text, element.options);
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: text 内容或 options 缺失`);
                }
                break;

            case "image":
                if (element.options && (element.options.path || element.options.data)) {
                    slide.addImage(element.options);
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: image 必须包含 path 或 data`);
                }
                break;

            case "table":
                if (element.options.rows && Array.isArray(element.options.rows)) {
                    slide.addTable(element.options.rows, element.options || {});
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: table 缺少有效 rows`);
                }
                break;

            case "chart":
                if (
                    element.options.chartType &&
                    element.options.data &&
                    Array.isArray(element.options.data)
                ) {
                    try {
                        const chartType = pres.ChartType[element.options.chartType.toLowerCase()];
                        console.log(element.options.chartType,"chartType")
                        if (chartType) {
                            slide.addChart(chartType, element.options.data, element.options || {});
                        } else {
                            console.warn(`无效的 chartType: ${element.options.chartType}`);
                        }
                    } catch (err) {
                        console.error(`添加图表失败:`, err);
                    }
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: chart 配置不完整`);
                }
                break;

            case "shape":
                if (element.options && element.options.shape) {
                    slide.addShape(element.options.shape, element.options);
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: shape 缺少 shapeName 或 options`);
                }
                break;

            case "media":
                if (element.options && element.options.path) {
                    slide.addMedia(element.options);
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: media 缺少 path`);
                }
                break;

            case "notes":
                if (element.text) {
                    slide.addNotes(element.text);
                } else {
                    console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: notes 缺少 text`);
                }
                break;
            case "slideNumber":
                slide.slideNumber(element.options || {});
                break;

            default:
                console.warn(`Slide ${slideIndex + 1}, Element ${elIndex + 1}: 未知类型 "${type}"，跳过渲染`);
        }
    });
});


pres.writeFile({ fileName: './ppt/bytemp2.pptx' });