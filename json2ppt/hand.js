import pptxgen from "pptxgenjs";
import { addMainTitle, addSpeechmaker, addSpeechTime, setBackgroundImage, setLogoImage } from "../tool/1-1.js";
// 1. Create a new Presentation
let pres = new pptxgen();

pres.defineSlideMaster({
    title: 'LOGO_TEMPLATE',
    // background: { fill: 'F2F2F2' }, // 背景浅灰色
    objects: [
        {
            image: {
                x: 0.5,
                y: 0.5,
                w: 0.5,
                h: 0.5,
                path: 'https://aiultronx.com/wp-content/uploads/2024/09/cropped-1-03.png', // 可替换为你的 logo 地址或 base64
            },
        },
        {
            text: '公司官网：www.yourcompany.com',
            options: {
                x: 1,
                y: 5.0,           // 👈 调整到底部区域（5.0 in ~ 5.625 in）
                w: '100%',
                h: 0.4,
                fontSize: 12,
                color: '666666',
                align: 'center',
            }
        },
    ],
})


// 第一页
// masterName 使用母版
let slide1 = pres.addSlide({ masterName: 'LOGO_TEMPLATE' })

setBackgroundImage(slide1, "C:Users/admin/Downloads/bg.png")
setLogoImage(slide1, "https://aiultronx.com/wp-content/uploads/2024/09/cropped-1-03.png")
addMainTitle(slide1, "深圳市爱奥创科技有限公司")
addSpeechmaker(slide1, "aoc")
addSpeechTime(slide1, "2025年7月19日")




// 第二页
let slide2 = pres.addSlide()
slide2.background = {
    color: 'FFF000', transparency: 50
}
// setBackgroundImage(slide2, "C:Users/admin/Downloads/bg.png")
// 目录
slide2.addText(`目录`, {
    x: 0,
    y: "10%",
    w: "100%",
    h: "15%",
    bold: true,
    fontSize: 22,
    // margin:10,
    align: "center",
    color: "000000", // pptx.SchemeColor.accent3,
});

slide2.addText(`第一条：员工培育计划`, {
    x: 4.5,
    y: "30%",
    w: "40%",
    h: "5%",

    fontSize: 14,
    fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
    // margin:10,
    align: "center",
    color: "000000", // pptx.SchemeColor.accent3,
});

slide2.addText(`第二条：员工职能分配`, {
    x: 4.5,
    y: "40%",
    w: "40%",
    h: "5%",

    fontSize: 14,
    fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
    // margin:10,
    align: "center",
    color: "000000", // pptx.SchemeColor.accent3,
});
slide2.addText(`第三条：奖惩机制`, {
    x: 4.5,
    y: "50%",
    w: "40%",
    h: "5%",

    fontSize: 14,
    fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
    // margin:10,
    align: "center",
    color: "000000", // pptx.SchemeColor.accent3,
});

slide2.addText(`第四条：汇报总结`, {
    x: 4.5,
    y: "60%",
    w: "40%",
    h: "5%",

    fontSize: 14,
    fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
    // margin:10,
    align: "center",
    color: "000000", // pptx.SchemeColor.accent3,
});

slide2.addText(`第五条：危机意识`, {
    x: 4.5,
    y: "70%",
    w: "40%",
    h: "5%",

    fontSize: 14,
    fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
    // margin:10,
    align: "center",
    color: "000000", // pptx.SchemeColor.accent3,
});

// 第三页
let slide3 = pres.addSlide()

slide3.addText("01", {
    x: "60%",
    y: "10%",
    h: "30%",
    w: "30%",
    fontSize: 100,
    bold: true,
    align: "center",
    color: "000000"
})

slide3.addText("员工培育计划", {
    x: "50%",
    y: "45%",
    h: 1,
    w: "40%",
    fontSize: 30,
    bold: true,
    align: "center",
    color: "000000"
})


let section1 = [
    {
        title: "销售额与订单量",
        desc: "本年度电商销售数据呈现出稳健增长的态势。销售额与订单量的同比增长显著，显示出电商市场需求的旺盛。特别是节假日促销期间，销售额出现爆发式增长，表明消费者对于电商购物模式的接受度和依赖度不断提高。"
    },
    {
        title: "热销品类与趋势",
        desc: "热销品类主要集中在电子产品、服装鞋帽、家居用品等领域。随着消费者对生活品质的追求，这些品类的销售额持续攀升。同时，绿色环保、健康养生等概念的商品也逐渐成为市场的新宠，显示出消费者对健康和环保的关注日益增强。"
    },
    {
        title: "用户增长与活跃度",
        desc: "电商平台的用户数量持续增长，新用户注册量不断攀升。用户活跃度也有所提高，用户粘性增强。这主要得益于电商平台不断优化用户体验，提供更加便捷、个性化的购物服务。"
    },
    {
        title: "区域销售差异",
        desc: "销售数据呈现出明显的区域差异。一线城市仍然是电商消费的主力军，销售额占比最高。但二线及以下城市的市场潜力巨大，销售额增长迅速，成为电商平台新的增长点。"
    }
]

// PPT正文，通常一个正文四个小点
let slide4 = pres.addSlide()


slide4.addText("员工培育计划", {
    x: 0.5,
    y: 0.5,
    w: "100%",
    h: "10%",
    bold: true,
    fontSize: 30
})

section1.forEach((item, index) => {
    let xrate = index * 20 + 10
    slide4.addText(item.title, {
        x: `${xrate}%`,
        y: "25%",
        w: "20%",
        h: "10%",
        bold: true,
        fontSize: 16
    })
    slide4.addText(item.desc, {
        x: `${xrate}%`,
        y: "35%",
        w: "20%",
        h: "30%",
        fontSize: 11
    })
})

let slide5 = pres.addSlide({ masterName: "LOGO_TEMPLATE" })
// masterName 使用母版

setBackgroundImage(slide5, "C:Users/admin/Downloads/bg.png")
setLogoImage(slide5, "https://aiultronx.com/wp-content/uploads/2024/09/cropped-1-03.png")
addMainTitle(slide5, "谢谢聆听！")
addSpeechmaker(slide5, "aoc")
addSpeechTime(slide5, "2025年7月19日")






pres.writeFile({ fileName: './ppt/CustomerReport1.pptx' });