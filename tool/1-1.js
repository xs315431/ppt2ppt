export function setLogoImage(slide, url) {
    slide.addImage({
        path: url,
        x: 0.5,
        y: 0.5,
        w: 0.5,
        h: 0.5
    });

}


export function setBackgroundImage(slide, url) {
    slide.background = {
        path: url
    }
}

export function addMainTitle(slide, title) {
    // 添加主题文字
    slide.addText(title, {
        x: 0.5,
        y: "40%",
        w: "100%",
        h: 0,
        fit: "resize",

        bold: true,
        fontSize: 40,
        align: "left",
        fill: { color: "FF0000" }, //{ color: pptx.SchemeColor.background2 },
        color: "000000", // pptx.SchemeColor.accent3,
    });
}

export function addEditTime(slide, str) {
    // 添加时间
    slide.addText(str, {
        x: 0,
        y: "90%",
        w: "100%",
        h: "10%",
        fontSize: 10,
        align: "left",
        color: "000000", // pptx.SchemeColor.accent3,
    });
}
export function addSpeechmaker(slide, Speechmaker) {
    // 添加演讲人
    slide.addText(`主讲人：${Speechmaker}`, {
        x: 0.5,
        y: "65%",
        w: "15%",
        h: "5%",
        line: {
            color: "FF0000",
            size: 1,
            dashType: "dashDot",
        },
        shape: "roundRect",
        rectRadius: 1,
        fontSize: 12,
        fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
        // margin:10,
        align: "left",
        color: "000000", // pptx.SchemeColor.accent3,
    });
}

export function addSpeechTime(slide, speechTime) {
    // 添加时间
    slide.addText(`时间:${speechTime}`, {
        x: 2.2,
        y: "65%",
        w: "17%",
        h: "5%",
        line: {
            color: "FF0000",
            size: 1,
            dashType: "dashDot",
        },
        shape: "roundRect",
        rectRadius: 1,
        fontSize: 12,
        fill: { color: "#eed7d7", transparency: 50 }, //{ color: pptx.SchemeColor.background2 },
        margin: [1, 2, 1, 2],
        align: "left",
        color: "000000",
    });
}



