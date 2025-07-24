# 生成ppt
- 技术：nodejs,python,pptgenjs,python-ppt

## 难点
- ppt模板文件手动编辑复杂度高 费时费力
  - pptx转为json文件：这里从定制的配置文件入手，通过配置文件的格式将pptx通过python-ppt转出，根据配置文件格式筛选出需要使用到的属性进行保存
  - pptgenjs通过json文件生成ppt：这个json文件必须是可以配合这两者进行做转换的
- ai生成的大纲如何替换json中的字符
- 该方案没有实时预览的功能
  - pptgenjs提供了通过nodejs生成代码并渲染导出的功能，没有办法实时预览，这里的一个解决方案是保存模板封面，用户根据封面进行主题的选择


## 进度
- ppt模板转为json文件
  - 目前适配了文字和图片信息的提取和转换，包括位置信息，大小信息，颜色
- 与json文件转为ppt文件
  - 适配了pptgenjs中元素: text,image,table,chart,shape,media,notes


1. 通过别人的ppt生成配置文件
2. 使用配置文件套如AI生成大纲文件
3. 使用pptgenjs生成ppt文件


ppt模板有了进展：我通过定义一种数据格式，将网上获取ppt模板转出一个配置文件，再使用pptgenjs通过配置json生成ppt


## 目录详情
- assets 存放起手的静态文件
- json2ppt nodejs文件模块通过配置文件生成ppt
- ppt2json python文件模块 将ppt模板转为json文件
- PPTTemplate ppt模板存放目录
- ppt ppt生成结果
- tool 工具函数存放