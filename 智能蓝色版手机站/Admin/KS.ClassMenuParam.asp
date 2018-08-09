<%
'全局参数设置
Const ChannelID="0"      '模块ID
Const MaxPerLine="10"     '每行显示数量
Const JsFileName="Menu.js"      '生成的JS文件名

'菜单显示参数设置
Const RCM_Menu_1="4"      '菜单弹出方式 1：左  2：右  3：上  4：下
Const RCM_Menu_2="0"      '菜单弹出横向偏移量
Const RCM_Menu_3="0"      '菜单弹出纵向偏移量
Const RCM_Menu_4="2"      '菜单项边距
Const RCM_Menu_5="3"      '菜单项间距
Const RCM_Menu_6="6"      '菜单项左边距
Const RCM_Menu_7="7"      '菜单项右边距
Const RCM_Menu_8="100"      '菜单透明度         0-100 完全透明-完全不透明
Const RCM_Menu_9="filter:Glow(Color=#000000, Strength=3)"      '其它特效
Const RCM_Menu_10="4"        '鼠标指在菜单项时，菜单弹出效果
Const RCM_Menu_11=""        '其它特效
Const RCM_Menu_12="23"        '鼠标移出菜单项时，菜单弹出效果
Const RCM_Menu_13="50"        '菜单弹出效果速度  10-100
Const RCM_Menu_14="2"        '弹出菜单阴影效果 0：none  1：simple  2：complex
Const RCM_Menu_15="4"        '弹出菜单阴影深度
Const RCM_Menu_16="#999999"        '弹出菜单阴影颜色
Const RCM_Menu_17="#ffffff"        '弹出菜单背景颜色
Const RCM_Menu_18=""        '弹出菜单背景图片，只有当菜单项背景颜色设为透明色：transparent 时才有效
Const RCM_Menu_19="3"        '弹出菜单背景图片平铺模式。 0：不平铺  1：横向平铺  2：纵向平铺  3：完全平铺
Const RCM_Menu_20="1"        '弹出菜单边框类型 0：无边框  1：单实线  2：双实线  5：凹陷  6：凸起
Const RCM_Menu_21="1"        '弹出菜单边框宽度
Const RCM_Menu_22="#ACA899"        '弹出菜单边框颜色
Const RCM_Menu_23="#ffffff"

'菜单项参数设置
Const RCM_Item_1="0"      '菜单项类型  0--Txt  1--Html  2--Image
Const RCM_Item_2=""       '菜单项名称
Const RCM_Item_3=""       '菜单项为Image，图片文件
Const RCM_Item_4=""       '菜单项为Image，鼠标指在菜单项时，图片文件。
Const RCM_Item_5="-1"     '菜单项为Image，图片宽度
Const RCM_Item_6="-1"     '菜单项为Image，图片高度
Const RCM_Item_7="0"      '菜单项为Image，图片边框
Const RCM_Item_8=""       '菜单项链接地址
Const RCM_Item_9=""       '菜单项链接目标 如：_self  _blank
Const RCM_Item_10=""      '菜单项链接状态栏显示
Const RCM_Item_11=""      '菜单项链接地址提示信息
Const RCM_Item_12=""        '菜单项左图片
Const RCM_Item_13=""        '鼠标指在菜单项时，菜单项左图片
Const RCM_Item_14="0"        '菜单项左图片宽度，0为图像文件原始值
Const RCM_Item_15="0"        '菜单项左图片高度，0为图像文件原始值
Const RCM_Item_16="0"        '菜单项左图片边框大小
Const RCM_Item_17=""        '菜单项右图片。如：arrow_r.gif
Const RCM_Item_18=""        '鼠标指在菜单项时，菜单项右图片。如：arrow_w.gif
Const RCM_Item_19="0"        '菜单项右图片宽度，0为图像文件原始值
Const RCM_Item_20="0"        '菜单项右图片高度，0为图像文件原始值
Const RCM_Item_21="0"        '菜单项右图片边框大小
Const RCM_Item_22="0"        '菜单项文字水平对齐方式  0：左对齐  1：居中  2：右对齐
Const RCM_Item_23="1"        '菜单项文字垂直对齐方式  0：顶部  1：居中  2：底部
Const RCM_Item_24="#F1F2EE"        '菜单项背景颜色  透明色：'transparent'
Const RCM_Item_25="1"        '菜单项背景颜色是否显示  0：显示  其它：不显示
Const RCM_Item_26="#CCCCCC"        '鼠标指在菜单项时，菜单项背景颜色
Const RCM_Item_27="1"        '鼠标指在菜单项时，菜单项背景颜色是否显示。  0：显示  其它：不显示
Const RCM_Item_28=""        '菜单项背景图片
Const RCM_Item_29=""        '鼠标指在菜单项时，菜单项背景图片
Const RCM_Item_30="3"        '菜单项背景图片平铺模式。 0：不平铺  1：横向平铺  2：纵向平铺  3：完全平铺
Const RCM_Item_31="3"     '鼠标指在菜单项时，菜单项背景图片平铺模式。0-3
Const RCM_Item_32="0"        '菜单项边框类型 0：无边框  1：单实线  2：双实线  5：凹陷  6：凸起
Const RCM_Item_33="0"        '菜单项边框宽度
Const RCM_Item_34="#FFFFF7"        '菜单项边框颜色
Const RCM_Item_35="#FF0000"        '鼠标指在菜单项时，菜单项边框颜色
Const RCM_Item_36="#000000"        '菜单项文字颜色
Const RCM_Item_37="#CC0000"        '鼠标指在菜单项时，菜单项文字颜色
Const FontSize_RCM_Item_38="9pt"        '菜单项文字大小
Const FontName_RCM_Item_38="宋体"        '菜单项文字字体
Const FontSize_RCM_Item_39="9pt"        '鼠标指在菜单项时,菜单项文字大小
Const FontName_RCM_Item_39="宋体"        '鼠标指在菜单项时,菜单项文字字体
%>