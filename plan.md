# python-pptx-ng 功能对齐 Open XML SDK 实施计划

## Context

Open-XML-SDK代码请参考/mnt/hdd1/pyopenxml/Open-XML-SDK
请一比一严格参考微软官方的SDK代码，避免引入错误。

python-pptx-ng 目前有 151 个 XML 元素类，覆盖了日常 PPT 操作的 ~80%。相比之下，C# Open XML SDK 有 1080 个 PresentationML 强类型类 + 133 个 OMML 数学公式类，覆盖了动画、切换、批注、SmartArt、数学公式等全部功能。本计划旨在分阶段补齐这些差距，同时建立验证框架确保质量。**允许破坏性 API 变更**。

**核心架构模式**（所有新功能必须遵循）：
1. `oxml/` 层：`BaseOxmlElement` 子类 + `register_element_cls()` 注册
2. `parts/` 层：`XmlPart` 子类 + `content_type_to_part_class_map` 注册
3. 高层 API：`ElementProxy`/`PartElementProxy` 子类
4. 枚举：`enum/` 模块

**关键注册文件**：
- `src/pptx/oxml/__init__.py` — XML 元素注册（~487 行）
- `src/pptx/__init__.py` — Content-Type → Part 映射（~83 行）
- `src/pptx/opc/constants.py` — 常量定义（已包含大部分 CT 和 RT 常量）

---

## Phase 1: 验证框架（高优先级）

### 1.1 Schema 验证引擎

**新建文件**：
- `src/pptx/validation/__init__.py` — 导出 `PresentationValidator`
- `src/pptx/validation/schema.py` — SchemaValidator 类
- `src/pptx/validation/context.py` — `ValidationErrorInfo`, `ValidationContext`
- `src/pptx/validation/package_validator.py` — 包结构验证

**修改文件**：
- `src/pptx/oxml/xmlchemy.py` — 在 `BaseOxmlElement` 上添加 `validate()` 方法，利用已有的 `ZeroOrOne`/`OneAndOnlyOne`/`RequiredAttribute` 等声明式约束
- `src/pptx/exc.py` — 添加 `ValidationError` 异常类
- `src/pptx/enum/validation.py` — `ValidationSeverity` 等枚举

**核心类**：
```
ValidationErrorInfo(error_type, id, description, part, element, xpath)
ValidationContext(errors: list[ValidationErrorInfo])
PresentationValidator
  .validate(presentation) -> list[ValidationErrorInfo]
  .validate_part(part) -> list[ValidationErrorInfo]
  .validate_element(element) -> list[ValidationErrorInfo]
SchemaValidator
  .validate_element(ctx, element)  # 检查子元素基数、属性值合法性
PackageValidator
  .validate(ctx, package)  # 检查必需 part、relationship 合法性
```

**验证逻辑要点**：
- 读取 `BaseOxmlElement._tag_seq` 验证子元素顺序
- 利用 `OptionalAttribute`/`RequiredAttribute` 元数据验证属性
- 委托 `simpletypes.py` 验证属性值
- 检查包结构（必需 part 如 presentation.xml、至少一个 slideMaster）

**复杂度**：XL

### 1.2 语义验证引擎

**新建文件**：
- `src/pptx/validation/semantic.py` — SemanticValidator + 约束基类
- `src/pptx/validation/constraints/__init__.py`
- `src/pptx/validation/constraints/attribute_constraints.py` — 属性约束（范围、配对、唯一性）
- `src/pptx/validation/constraints/reference_constraints.py` — 引用完整性（rId 指向有效 part）
- `src/pptx/validation/constraints/relationship_constraints.py` — relationship 类型约束

**核心约束**（参考 C# SDK 的 20 种语义约束类型）：
- `AttributeCannotOmitConstraint` — 必需属性
- `AttributeValueRangeConstraint` — 值范围（如 sldId ∈ [256, 2^31]）
- `UniqueAttributeValueConstraint` — 唯一性（如 cNvPr/@id 在 slide 内唯一）
- `ReferenceExistConstraint` — rId 指向存在的 part
- `AttributePairConstraint` — 属性共存条件

**复杂度**：L

---

## Phase 2: 核心演示功能

### 2.1 批注系统

**新建文件**：
- `src/pptx/oxml/comment.py` — CT_CommentAuthorList, CT_CommentAuthor, CT_CommentList, CT_Comment
- `src/pptx/comment.py` — CommentAuthors, CommentAuthor, Comments, Comment 高层代理
- `src/pptx/parts/comment.py` — CommentAuthorsPart, CommentPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 `p:cmAuthorLst`, `p:cmAuthor`, `p:cmLst`, `p:cm`
- `src/pptx/__init__.py` — 添加 CT 映射：`CT.PML_COMMENT_AUTHORS → CommentAuthorsPart`, `CT.PML_COMMENTS → CommentPart`（常量已存在于 constants.py）
- `src/pptx/parts/slide.py` — 添加 `SlidePart.comment_part` 属性
- `src/pptx/parts/presentation.py` — 添加 `PresentationPart.comment_authors_part` 属性
- `src/pptx/slide.py` — 添加 `Slide.comments` 属性

**高层 API**：
```python
prs = Presentation()
slide = prs.slides[0]
author = prs.comment_authors.add("张三", "ZS")
comment = slide.comments.add("这是一条批注", author, left=Emu(914400), top=Emu(914400))
```

**XML 元素**：`p:cmAuthorLst`, `p:cmAuthor`, `p:cmLst`, `p:cm`
**复杂度**：M

### 2.2 演示属性和视图属性

**新建文件**：
- `src/pptx/oxml/presprops.py` — CT_PresentationProperties（含 showPr, prnPr 等子元素）
- `src/pptx/oxml/viewprops.py` — CT_ViewProperties（含 normalViewPr, slideViewPr 等）
- `src/pptx/presprops.py` — PresentationProperties, ShowProperties 代理
- `src/pptx/viewprops.py` — ViewProperties 代理
- `src/pptx/parts/presprops.py` — PresentationPropertiesPart
- `src/pptx/parts/viewprops.py` — ViewPropertiesPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 `p:presentationPr`, `p:viewPr` 等
- `src/pptx/__init__.py` — 添加 CT 映射（常量已存在）
- `src/pptx/parts/presentation.py` — 添加 pres_props_part, view_props_part 属性
- `src/pptx/presentation.py` — 暴露 `Presentation.properties`, `Presentation.view_properties`

**高层 API**：
```python
prs = Presentation()
prs.properties.show_mode = PP_SHOW_MODE.KIOSK
prs.properties.loop_continuously = True
prs.view_properties.show_grid = True
```

**XML 元素**：`p:presentationPr`, `p:showPr`, `p:prnPr`, `p:viewPr`, `p:normalViewPr`, `p:slideViewPr`, `p:gridSpacing`
**复杂度**：M

### 2.3 自定义放映

**新建文件**：
- `src/pptx/oxml/customshow.py` — CT_CustomShowList, CT_CustomShow, CT_SlideList, CT_SlideListEntry
- `src/pptx/customshow.py` — CustomShows, CustomShow 代理

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 `p:custShowLst`, `p:custShow`, `p:sldLst`
- `src/pptx/oxml/presentation.py` — CT_Presentation 添加 `custShowLst` 子元素
- `src/pptx/presentation.py` — 添加 `Presentation.custom_shows` 属性

**复杂度**：S

### 2.4 幻灯片切换效果

**新建文件**：
- `src/pptx/oxml/transition.py` — CT_Transition + 13 种具体切换类型元素类
- `src/pptx/transition.py` — Transition 代理类
- `src/pptx/enum/transition.py` — PP_TRANSITION_TYPE, PP_TRANSITION_SPEED, PP_DIRECTION 枚举

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 22 个元素（见下）
- `src/pptx/oxml/slide.py` — CT_Slide 已有 transition 的 ZeroOrOne 声明，需增强
- `src/pptx/slide.py` — 添加 `Slide.transition` 属性

**XML 元素**（22 个）：
`p:transition`, `p:blinds`, `p:checker`, `p:comb`, `p:cover`, `p:cut`, `p:dissolve`, `p:fade`, `p:pull`, `p:push`, `p:random`, `p:randomBar`, `p:split`, `p:strips`, `p:wedge`, `p:wheel`, `p:wipe`, `p:zoom`, `p:circle`, `p:diamond`, `p:newsflash`, `p:plus`

**高层 API**：
```python
slide = prs.slides[0]
slide.transition = Transition()
slide.transition.set_fade()
slide.transition.speed = PP_TRANSITION_SPEED.FAST
slide.transition.advance_after = 5000  # ms
```

**复杂度**：M

### 2.5 动画系统

**新建文件**：
- `src/pptx/oxml/animation.py` — ~20 个 CT_* 动画元素类
- `src/pptx/animation.py` — AnimationTimeline, AnimationEffect 代理
- `src/pptx/enum/animation.py` — PP_ANIMATION_TYPE, PP_TIMING_NODE_TYPE 等枚举

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册动画元素
- `src/pptx/oxml/slide.py` — 增强 CT_SlideTiming（当前只有基础定义）
- `src/pptx/slide.py` — 添加 `Slide.timing` 属性

**核心 XML 元素**（~20 个）：
`p:timing`（已有）, `p:tnLst`（已有）, `p:childTnLst`（已有）, `p:par`, `p:seq`, `p:excl`, `p:cTn`, `p:anim`, `p:animClr`, `p:animEffect`, `p:animMotion`, `p:animRot`, `p:animScale`, `p:set`, `p:cmd`, `p:cBhvr`, `p:tgtEl`, `p:spTgt`, `p:cond`, `p:stCondLst`, `p:tav`, `p:tavLst`, `p:bldLst`, `p:bldP`, `p:cMediaNode`

**复杂度**：XL（最复杂的单一功能）

---

## Phase 3: 媒体与高级内容

### 3.1 增强音视频支持

**新建文件**：
- `src/pptx/audio.py` — Audio 值对象

**修改文件**：
- `src/pptx/media.py` — 添加 Audio 类
- `src/pptx/parts/slide.py` — 添加 `get_or_add_audio_media_part()` 方法
- `src/pptx/shapes/shapetree.py` — 添加 `add_audio()` 方法，增强 `add_video()`
- `src/pptx/parts/media.py` — 添加音频 content type 支持
- `src/pptx/oxml/slide.py` — 增强媒体节点类（p:cMediaNode 属性）

**复杂度**：M

### 3.2 讲义母版

**新建文件**：
- `src/pptx/oxml/handout.py` — CT_HandoutMaster
- `src/pptx/handout.py` — HandoutMaster 代理
- `src/pptx/parts/handout.py` — HandoutMasterPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 `p:handoutMaster`
- `src/pptx/__init__.py` — 添加 CT 映射（常量已存在）
- `src/pptx/parts/presentation.py` — 添加 handout_master_part 属性

**复杂度**：S

### 3.3 用户自定义标签

**新建文件**：
- `src/pptx/oxml/tags.py` — CT_TagList, CT_Tag
- `src/pptx/tags.py` — UserDefinedTags 代理
- `src/pptx/parts/tags.py` — UserDefinedTagsPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 `p:tagLst`, `p:tag`
- `src/pptx/__init__.py` — 添加 CT 映射（常量已存在）
- `src/pptx/parts/slide.py` — 添加 tags_part 属性
- `src/pptx/slide.py` — 添加 `Slide.tags` 属性

**复杂度**：S

### 3.4 SmartArt/图示支持

**新建文件**：
- `src/pptx/oxml/diagram/__init__.py`
- `src/pptx/oxml/diagram/data.py` — CT_DiagramData, CT_PointList, CT_Point, CT_TextLink 等
- `src/pptx/oxml/diagram/layout.py` — CT_DiagramDefinition, CT_LayoutNode 等
- `src/pptx/oxml/diagram/style.py` — CT_DiagramStyle
- `src/pptx/oxml/diagram/colors.py` — CT_DiagramColors
- `src/pptx/diagram.py` — SmartArtDiagram 代理
- `src/pptx/parts/diagram.py` — DiagramDataPart, DiagramLayoutPart, DiagramStylePart, DiagramColorsPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 diagram 元素
- `src/pptx/__init__.py` — 添加 diagram CT 映射（常量已存在）
- `src/pptx/oxml/ns.py` — 添加 `dgm` 命名空间前缀
- `src/pptx/shapes/shapetree.py` — 添加 `add_smartart()` 方法

**复杂度**：XL（4 个相互关联的 XML part）

### 3.5 扩展图表（chartEx）

**新建文件**：
- `src/pptx/oxml/chartex/__init__.py`
- `src/pptx/oxml/chartex/chartex.py` — chartEx 命名空间的元素类
- `src/pptx/chartex.py` — 扩展图表代理
- `src/pptx/parts/chartex.py` — ChartExPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 chartEx 元素
- `src/pptx/__init__.py` — 添加 CT 映射（CT.OFC_CHART_EX 常量已存在）
- `src/pptx/oxml/ns.py` — 添加 `cx` 命名空间前缀

**复杂度**：L

### 3.6 3D 模型支持

**新建文件**：
- `src/pptx/oxml/model3d.py` — 3D 模型元素类

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 3D 模型元素
- `src/pptx/oxml/ns.py` — 添加 `m3d` 命名空间前缀

**复杂度**：S

### 3.7 OMML 数学公式（Office Math Markup Language）

**目标**：支持在幻灯片中创建、读取和修改数学公式（Office 公式编辑器格式）。

**现状**：python-pptx-ng 仅在 `ns.py` 中定义了 `"m"` 命名空间前缀（`http://schemas.openxmlformats.org/officeDocument/2006/math`），但没有任何 OMML 元素类（0 个）。C# Open XML SDK 有 **133 个** OMML 强类型类。

**公式在 PPT 中的嵌入方式**：
- 段落公式：`<m:oMathPara>` 作为文本段落中的独立数学段落
- 行内公式：`<m:oMath>` 嵌套在 `<a:p>` 内，与其他 `<a:r>` 文本并列
- 公式内容使用 OMML 标记（`m:` 命名空间）

**参考 C# SDK 核心类**（133 个，按功能分组）：

| 功能 | 类名 | XML 元素 |
|------|------|----------|
| 公式容器 | `oMath`, `oMathPara` | `m:oMath`, `m:oMathPara` |
| 分数 | `Fraction`, `FractionProperties`, `FractionType` | `m:f`, `m:fPr` |
| 根号 | `Rad` | `m:rad` |
| 上标 | `SuperArgument` | `m:sSup` |
| 下标 | `SubArgument` | `m:sSub` |
| 上下标 | `SubSuperArgument` | `m:sSubSup` |
| N 元运算符（积分/求和） | `Nary`, `NaryProperties`, `NaryLimitLocation` | `m:nary`, `m:naryPr` |
| 矩阵 | `Matrix`, `MatrixRow`, `MatrixColumn`, `MatrixColumns` | `m:m`, `m:mr`, `m:mc` |
| 函数 | `MathFunction`, `FunctionName`, `FunctionProperties` | `m:func`, `m:fName` |
| 分隔符（括号） | `Delimiter`, `DelimiterProperties` | `m:d`, `m:dPr` |
| 重音符号 | `Accent`, `AccentProperties` | `m:acc`, `m:accPr` |
| 上下划线 | `Bar`, `BarProperties` | `m:bar`, `m:barPr` |
| 边框盒 | `BorderBox`, `BorderBoxProperties` | `m:borderBox`, `m:borderBoxPr` |
| 方程数组 | `EquationArray`, `EquationArrayProperties` | `m:eqArr`, `m:eqArrPr` |
| 分组字符 | `GroupChar`, `GroupCharProperties` | `m:groupChr`, `m:groupChrPr` |
| 上下限 | `LimitLower`, `LimitUpper`, `Limit` | `m:limLow`, `m:limUpp` |
| 运算符盒子 | `Box`, `BoxProperties` | `m:box`, `m:boxPr` |
| 通用子元素 | `Run` (文本), `ArgumentProperties`, `ControlProperties` | `m:r`, `m:rPr`, `m:ctrlPr` |
| 全局属性 | `MathProperties`, `MathFont`, `DisplayDefaults` | `m:mathPr`, `m:mathFont` |
| 结构辅助 | `Base`, `Degree`, `Numerator`, `Denominator`, `Superscript`, `Subscript` | `m:e`, `m:deg`, `m:num`, `m:den`, `m:sup`, `m:sub` |

**新建文件**：
- `src/pptx/oxml/math/__init__.py` — 包初始化
- `src/pptx/oxml/math/omath.py` — 核心公式元素类（CT_OMath, CT_OMathPara, CT_OMathRun 等）
- `src/pptx/oxml/math/fraction.py` — 分数元素类（CT_OMathFrac, CT_OMathFracProperties）
- `src/pptx/oxml/math/radical.py` — 根号元素类（CT_OMathRad, CT_OMathRadProperties）
- `src/pptx/oxml/math/scripts.py` — 上下标元素类（CT_OMathSup, CT_OMathSub, CT_OMathSubSup）
- `src/pptx/oxml/math/nary.py` — N 元运算符（CT_OMathNary — 积分 ∫、求和 Σ、乘积 Π 等）
- `src/pptx/oxml/math/matrix.py` — 矩阵元素类（CT_OMathMatrix, CT_OMathMatrixRow, CT_OMathMatrixColumn）
- `src/pptx/oxml/math/delimiter.py` — 分隔符/括号（CT_OMathDelimiter）
- `src/pptx/oxml/math/accent.py` — 重音符号、上下划线（CT_OMathAccent, CT_OMathBar）
- `src/pptx/oxml/math/equation.py` — 方程数组、边框盒、分组字符、上下限
- `src/pptx/oxml/math/properties.py` — 公式属性元素（CT_OMathMathPr, CT_DisplayDefaults 等）
- `src/pptx/math.py` — 高层公式代理类（MathFormula, MathFraction, MathRadical 等）
- `src/pptx/enum/math.py` — 数学公式枚举（PP_FRACTION_TYPE, PP_NARY_TYPE, PP_LIMIT_LOCATION 等）

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册所有 `m:*` 元素（约 50 个）
- `src/pptx/oxml/text.py` — CT_TextParagraph 支持 `m:oMathPara` 和 `m:oMath` 作为子元素
- `src/pptx/text/text.py` — 添加 `_Paragraph.add_math()` 方法
- `src/pptx/text/fonts.py` — 支持 `m:mathFont` 字体属性

**高层 API**：
```python
from pptx.math import MathFormula, MathFraction, MathRadical, MathNary
from pptx.enum.math import PP_NARY_TYPE, PP_FRACTION_TYPE

slide = prs.slides[0]
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame

# 构建公式
formula = MathFormula()
frac = formula.add_fraction()  # 默认分数线型
frac.numerator.add_text("a")
frac.denominator.add_text("b")
integral = formula.add_nary(PP_NARY_TYPE.INTEGRAL)
integral.sub.add_text("0")
integral.sup.add_text("∞")
integral.base.add_fraction()  # 嵌套

p = tf.paragraphs[0]
p.add_math(formula)

# 读取已有公式
for para in tf.paragraphs:
    if para.has_math:
        math_element = para.math
        # 遍历公式结构树
```

**核心 XML 元素**（约 50 个需注册）：
```
m:oMath, m:oMathPara, m:r, m:rPr, m:t,                      # 基础容器与文本
m:f, m:fPr, m:num, m:den, m:fType,                           # 分数
m:rad, m:radPr, m:deg, m:e,                                  # 根号
m:sSup, m:sSupPr, m:sup,                                     # 上标
m:sSub, m:sSubPr, m:sub,                                     # 下标
m:sSubSup, m:sSubSupPr,                                      # 上下标
m:nary, m:naryPr, m:limLoc, m:chr,                           # N 元运算符（积分/求和）
m:m, m:mPr, m:mr, m:mc, m:mcPr,                              # 矩阵
m:d, m:dPr, m:begChr, m:endChr,                              # 分隔符（括号）
m:acc, m:accPr, m:bar, m:barPr,                              # 重音/划线
m:borderBox, m:borderBoxPr, m:box, m:boxPr,                  # 边框/盒子
m:eqArr, m:eqArrPr,                                           # 方程数组
m:func, m:funcPr, m:fName,                                   # 函数
m:groupChr, m:groupChrPr,                                    # 分组字符
m:limLow, m:limLowPr, m:limUpp, m:limUppPr, m:lim,           # 上下限
m:mathPr, m:mathFont, m:ctrlPr, m:defaultJust,               # 全局属性
```

**命名空间**：
- `"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"` — **已存在于 `ns.py`**
- `"mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"` — 兼容性标记（可能需要新增）

**复杂度**：XL（133 个 C# 类，约 50 个关键 XML 元素，公式结构深度嵌套）

---

## Phase 4: 打包与高级功能

### 4.1 表格样式

**新建文件**：
- `src/pptx/oxml/tablestyle.py` — CT_TableStyles, CT_TableStyle 元素类
- `src/pptx/tablestyle.py` — TableStyles 代理
- `src/pptx/parts/tablestyle.py` — TableStylesPart

**修改文件**：
- `src/pptx/oxml/__init__.py` — 注册 `p:tblStyleLst`, `p:tblStyle`
- `src/pptx/__init__.py` — 添加 CT 映射（常量已存在）

**复杂度**：M

### 4.2 主题深度操作

**修改文件**：
- `src/pptx/oxml/theme.py` — 添加子元素类：CT_ColorScheme, CT_FontScheme, CT_FormatScheme 等
- 新建 `src/pptx/theme.py` — Theme, ColorScheme, FontScheme 代理

**复杂度**：L

### 4.3 VBA/宏项目支持

**新建文件**：
- `src/pptx/parts/vba.py` — VbaProjectPart（二进制 passthrough）

**修改文件**：
- `src/pptx/__init__.py` — 确保 CT.PML_PRES_MACRO_MAIN 映射正确
- `src/pptx/api.py` — 处理 .pptm 文件检测

**复杂度**：S

### 4.4 幻灯片同步数据

**新建文件**：
- `src/pptx/parts/slidesync.py` — SlideSyncDataPart

**修改文件**：
- `src/pptx/__init__.py` — 添加 CT 映射（常量已存在）

**复杂度**：S

### 4.5 数字签名

**新建文件**：
- `src/pptx/dsignature.py` — 数字签名相关

**复杂度**：M

### 4.6 Ribbon/UI 自定义

**新建文件**：
- `src/pptx/oxml/customui.py` — customUI 元素解析

**复杂度**：M

### 4.7 ActiveX 控件

**新建文件**：
- `src/pptx/oxml/control.py` — CT_Control, CT_ControlList

**复杂度**：M

### 4.8 VML 绘图

**复杂度**：L（遗留格式，仅做 round-trip）

### 4.9 Web 扩展

**新建文件**：
- `src/pptx/parts/webextension.py` — WebExtensionPart

**复杂度**：M

---

## 跨切面修改

### 命名空间新增（`src/pptx/oxml/ns.py`）
```python
"p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"
"p15": "http://schemas.microsoft.com/office/powerpoint/2012/main"
"dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram"
"cx":  "http://schemas.microsoft.com/office/drawing/2014/chartex"
"m3d": "http://schemas.microsoft.com/office/drawing/2017/model3d"
"mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006"  # OMML 兼容性标记
```
注：`"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"` 已存在。

### 每个功能的标准注册流程
1. `oxml/` 新模块定义 CT_* 类
2. `oxml/__init__.py` 导入并 `register_element_cls()`
3. `parts/` 新模块定义 Part 类
4. `__init__.py` 添加 CT → Part 映射
5. 高层模块提供用户友好 API
6. `enum/` 模块提供枚举值

---

## 各阶段建议实施顺序

**Phase 1**：1.1 → 1.2（验证框架是后续所有功能的测试基础）

**Phase 2**：2.1 → 2.2 → 2.3 → 2.4 → 2.5
（批注 → 属性 → 自定义放映 → 切换 → 动画，从简单到复杂）

**Phase 3**：3.1 → 3.2 → 3.3 → 3.6 → 3.7 → 3.4 → 3.5
（音视频 → 讲义 → 标签 → 3D模型 → **数学公式** → SmartArt → 扩展图表，先做小功能）

**Phase 4**：4.3 → 4.4 → 4.1 → 4.2 → 4.5 → 4.6 → 4.7 → 4.9 → 4.8
（按实用价值排序）

---

## 验证方案

每个功能完成后：
1. 运行 Phase 1 的验证器对生成的 .pptx 文件进行 Schema + 语义验证
2. 用 PowerPoint/WPS 打开生成文件验证可视正确性
3. 往返测试：读取 → 修改 → 保存 → 再读取，验证数据完整性
4. 单元测试覆盖每个 CT_* 类的属性读写和子元素操作
5. 集成测试覆盖高层 API 的端到端场景

**运行测试**：`python -m pytest tests/ -v`
**验证文件**：用 `PresentationValidator.validate()` 检查生成的 .pptx
