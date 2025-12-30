# PPTX XML 高频核心标签与属性详解（含动画/表格，Markdown 表格版）

### 文档说明

- 命名空间前缀：`p:`=PresentationML（核心）、`a:`=DrawingML（样式/几何）、`r:`=Relationships（资源引用）、`s:`=SharedML（共享样式）、`anim:`=AnimationML（动画）

- 单位：坐标/尺寸均用 EMU（1cm=360000 EMU，1pt=12700 EMU）

- 空元素：`CT_Empty` 类型元素无内容/属性，仅作标记（如 `<p:honeycomb/>`）

- 适用场景：XML 解析、PPTX 二次开发、数据提取

---

## 模块 1：演示文稿全局配置（对应文件：ppt/presentation.xml）

负责演示文稿整体设置、幻灯片清单、尺寸等核心配置

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|presentation|p:|showComments|是否显示批注（1=显示，0=隐藏）|<p:presentation showComments="1"/>|
|presentation|p:|defaultTextStyle|全局默认文本样式引用（关联样式定义）|<p:presentation defaultTextStyle r:id="rId1"/>|
|sldIdLst|p:|-|幻灯片 ID 列表容器（存储所有幻灯片的标识）|<p:sldIdLst><p:sldId/></p:sldIdLst>|
|sldId|p:|id|幻灯片唯一标识（数字，全局不重复）|<p:sldId id="256" r:id="rId3"/>|
|sldId|p:|r:id|关联 slides/slideX.xml 的关系 ID（定位具体幻灯片文件）|<p:sldId id="256" r:id="rId3"/>|
|sldSz|p:|cx/cy|幻灯片宽/高（单位：EMU）|<p:sldSz cx="9144000" cy="6858000"/>|
|sldSz|p:|type|幻灯片类型（screen=屏幕展示，print=打印）|<p:sldSz type="screen"/>|
---

## 模块 2：单张幻灯片核心（对应文件：ppt/slides/slideX.xml）

单页幻灯片的根节点与画布配置，包含所有形状、文本、图片等元素

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|slide|p:|showMasterSp|是否显示母版形状（1=显示，继承母版样式；0=隐藏）|<p:slide showMasterSp="1"/>|
|slide|p:|showPh|是否显示占位符（1=显示，0=隐藏）|<p:slide showPh="1"/>|
|cSld|p:|-|幻灯片画布容器（必选，所有可视元素的父容器）|<p:cSld><p:spTree/></p:cSld>|
|spTree|p:|-|形状树（所有形状、组合图形、图片的根容器）|<p:spTree><p:sp/><p:grpSp/></p:spTree>|
|slideLayoutId|p:|val|幻灯片布局 ID（1=标题页，2=标题+内容，3=节标题，4=空白页等）|<p:slideLayoutId val="2"/>|
---

## 模块 3：组合图形（Group Shape）

多个形状/图片的组合容器，核心标识为 <p:grpSp>，区别于单个形状的 <p:sp>

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|grpSp|p:|-|组合图形根节点（唯一标识组合图形）|<p:grpSp><p:nvGrpSpPr/></p:grpSp>|
|nvGrpSpPr|p:|-|组合图形非视觉属性容器（存储 ID、名称等非渲染信息）|<p:nvGrpSpPr><p:cNvPr/></p:nvGrpSpPr>|
|cNvPr|p:|id/name|组合图形唯一 ID/名称（默认命名为 Group X，X 为数字）|<p:cNvPr id="10" name="Group 1"/>|
|grpSpPr|p:|-|组合图形视觉属性容器（存储位置、尺寸、样式等渲染信息）|<p:grpSpPr><a:xfrm/></p:grpSpPr>|
|xfrm|a:|off x/off y|组合图形左上角相对于幻灯片的坐标（单位：EMU）|<a:xfrm><a:off x="500000" y="800000"/></a:xfrm>|
|xfrm|a:|ext cx/ext cy|组合图形的整体宽/高（单位：EMU）|<a:xfrm><a:ext cx="6000000" cy="4000000"/></a:xfrm>|
|xfrm|a:|chOff/chExt|子形状相对于组合图形的偏移/可用区域（默认 x=0,y=0，与组合宽高一致）|<a:xfrm><a:chOff x="0" y="0"/<a:chExt cx="6000000" cy="4000000"/></a:xfrm>|
---

## 模块 4：单个形状（Single Shape）

独立形状（矩形、文本框、蜂窝形等）的核心结构，根标签为 <p:sp>

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|sp|p:|id/name|形状唯一 ID/名称（默认按形状类型命名，如 Rectangle 1）|<p:sp id="11" name="Rectangle 1"/>|
|nvSpPr|p:|-|形状非视觉属性容器（存储 ID、名称、描述等）|<p:nvSpPr><p:cNvPr/></p:nvSpPr>|
|spPr|p:|-|形状视觉属性容器（存储位置、尺寸、形状类型、样式等）|<p:spPr><a:prstGeom/></p:spPr>|
|prstGeom|a:|prst|预设形状类型（rect=矩形，ellipse=椭圆，honeycomb=蜂窝形，star=五角星等）|<a:prstGeom prst="honeycomb"/>|
|custGeom|a:|-|自定义形状容器（存储自定义路径，区别于预设形状）|<a:custGeom><a:pathLst/></a:custGeom>|
---

## 模块 5：文本内容与样式（对应容器：txBody）

形状内文本的容器、段落格式、字符样式等配置

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|txBody|p:|-|文本容器（所有文本相关元素的父容器）|<p:txBody><a:bodyPr/><a:p/></p:txBody>|
|bodyPr|a:|vertOverflow|文本垂直溢出处理（clip=裁剪，ellipsis=省略号，overflow=溢出显示）|<a:bodyPr vertOverflow="ellipsis"/>|
|bodyPr|a:|anchor|文本在容器内的对齐方式（t=上对齐，ctr=居中，b=下对齐，just=两端对齐）|<a:bodyPr anchor="ctr"/>|
|p|a:|algn|段落对齐方式（l=左对齐，r=右对齐，ctr=居中，just=两端对齐）|<a:p algn="ctr"/>|
|r|a:|-|同格式文本段（一段连续的相同样式文本）|<a:r><a:rPr/><a:t/></a:r>|
|rPr|a:|sz|字号（单位：1/100 点，如 2400=24pt）|<a:rPr sz="2400" b="1"/>|
|rPr|a:|b/i|加粗/斜体（1=启用，0=禁用）|<a:rPr sz="2400" b="1" i="0"/>|
|rPr|a:|color|文本颜色（子元素 srgbClr 定义具体颜色值）|<a:rPr><a:color><a:srgbClr val="#FF0000"/></a:color></a:rPr>|
|t|a:|-|纯文本内容（PPT 中显示的实际文本）|<a:t>演示文稿标题</a:t>|
---

## 模块 6：图片元素（对应标签：pic）

图片容器与资源引用，关联 media 目录下的实际图片文件

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|pic|p:|-|图片根节点（唯一标识图片元素）|<p:pic><p:nvPicPr/><p:blipFill/></p:pic>|
|nvPicPr|p:|cNvPr id/name|图片唯一 ID/名称（默认命名为 Picture X）|<p:nvPicPr><p:cNvPr id="12" name="Picture 3"/></p:nvPicPr>|
|blipFill|a:|rotWithShape|图片是否随形状旋转（1=是，0=否）|<a:blipFill rotWithShape="1"/>|
|blip|a:|r:embed|图片资源关系 ID（关联 .rels 文件中的图片路径）|<a:blip r:embed="rId2"/>|
|srcRect|a:|l/t/r/b|图片裁剪偏移（左/上/右/下，单位：EMU，正数为裁剪）|<a:srcRect l="1000" t="1000" r="2000" b="2000"/>|
---

## 模块 7：资源引用（对应文件：.rels）

管理幻灯片与外部资源（图片、媒体、其他幻灯片、主题）的关联关系，文件路径：ppt/slides/_rels/slideX.xml.rels

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|Relationship|r:|Id|关系 ID（与 r:embed、r:id 等属性对应，唯一标识资源）|<Relationship Id="rId2" Type="..." Target="..."/>|
|Relationship|r:|Type|资源类型（标识资源类别，如图片、幻灯片、主题）|Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"|
|Relationship|r:|Target|资源路径（相对路径，指向实际资源文件）|Target="../media/image1.png"|
---

## 模块 8：母版与布局（对应文件：slideMaster/slideLayout）

控制演示文稿全局样式与幻灯片布局，文件路径：ppt/slideMasters/slideMasterX.xml、ppt/slideLayouts/slideLayoutX.xml

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|sldMaster|p:|preserve|是否保留母版（1=保留，0=不保留，保留后可继承样式）|<p:sldMaster preserve="1"/>|
|sldLayout|p:|type|布局类型（title=标题页，textOnly=纯文本页，blank=空白页，chart=图表页等）|<p:sldLayout type="title"/>|
|ph|p:|type|占位符类型（title=标题占位符，body=内容占位符，chart=图表占位符等）|<p:ph type="title" idx="0"/>|
|ph|p:|idx|占位符索引（同一页面内唯一，用于区分多个占位符）|<p:ph type="body" idx="1"/>|
---

## 模块 9：动画标签（AnimationML，对应文件：slideX.xml/animations.xml）

控制幻灯片元素的动画效果，包含动画触发、持续时间、动画类型等配置

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|animations|p:|-|动画列表容器（存储单张幻灯片的所有动画）|<p:animations><p:anim/></p:animations>|
|anim|p:|spid|动画作用对象的形状 ID（关联具体形状/元素）|<p:anim spid="11" type="entrance"/>|
|anim|p:|type|动画类型（entrance=进入动画，exit=退出动画，emphasis=强调动画，motion=运动动画）|<p:anim spid="11" type="entrance"/>|
|anim|p:|dur|动画持续时间（单位：s=秒，如 2s=2秒）|<p:anim spid="11" dur="2s"/>|
|anim|p:|delay|动画延迟时间（单位：s=秒，如 0.5s=延迟0.5秒触发）|<p:anim spid="11" delay="0.5s"/>|
|custAnim|p:|-|自定义动画容器（存储自定义动画的详细配置）|<p:custAnim><anim:tm/></p:custAnim>|
|tm|anim:|dur|自定义动画持续时间（单位：ms=毫秒，如 1500=1.5秒）|<anim:tm dur="1500"/>|
|trigger|p:|type|动画触发方式（onClick=点击触发，onNext=下一页触发，onPrev=上一页触发，withPrevious=与前一个同时触发）|<p:trigger type="onClick"/>|
---

## 模块 10：表格标签（对应文件：slideX.xml）

幻灯片中表格的结构、单元格样式、内容配置

|标签|前缀|核心属性|属性含义|示例|
|---|---|---|---|---|
|tbl|a:|-|表格根节点（唯一标识表格元素）|<a:tbl><a:tblPr/><a:tr/></a:tbl>|
|tblPr|a:|firstRow|是否设置首行样式（1=是，0=否，用于区分表头）|<a:tblPr firstRow="1"/>|
|tblPr|a:|bandRow|是否设置条纹行样式（1=是，0=否，用于交替行颜色）|<a:tblPr bandRow="1"/>|
|tblGrid|a:|-|表格网格容器（定义表格列数与列宽）|<a:tblGrid><a:gridCol/></a:tblGrid>|
|gridCol|a:|w|表格列宽（单位：EMU）|<a:gridCol w="2000000"/>|
|tr|a:|h|表格行高（单位：EMU）|<a:tr h="500000"/>|
|tc|a:|gridSpan|单元格跨列数（如 2=跨2列）|<a:tc gridSpan="2"/>|
|tc|a:|rowSpan|单元格跨行数（如 2=跨2行）|<a:tc rowSpan="2"/>|
|tcPr|a:|fill|单元格填充样式（子元素定义填充颜色/图案）|<a:tcPr><a:fill><a:srgbClr val="#F5F5F5"/></a:fill></a:tcPr>|
|txBody|a:|-|单元格文本容器（存储单元格内的文本与样式）|<a:txBody><a:p><a:t>单元格内容</a:t></a:p></a:txBody>|
---

### 实操要点补充

1. 组合图形识别：根标签 <p:grpSp> + 子元素含 <sp>/<pic>/<grpSp> + 非视觉属性 <nvGrpSpPr>

2. 资源定位：通过 <Relationship> 的 Id 关联 r:embed/r:id，再通过 Target 路径找到实际资源（如图片）

3. 单位换算：EMU→cm 除以 360000，EMU→pt 除以 12700，EMU→px（96 DPI）除以 9525

4. 全量标签获取：下载官方 XSD 文件（[https://schemas.openxmlformats.org/](https://schemas.openxmlformats.org/)），用 VS Code XML Tools 插件生成带注释的全量文档
> （注：文档部分内容可能由 AI 生成）