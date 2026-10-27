# 图片放大预览 · Word / Excel / PowerPoint 插件（WPS 同款体验）

在 Word、Excel、PowerPoint 里，**单击图片 → 图片旁浮出放大镜按钮 → 点击 → 以原图尺寸预览**，
预览窗里滚轮缩放、左键拖动平移。图片在文档里被缩到 1cm，也能看到原始高清像素。
三个功能（放大镜预览 / 裁剪 / 旋转）在三个宿主里行为一致。

- 一个 DLL，同时挂载到 **Word**、**Excel**、**PowerPoint**
- **只服务 Microsoft Office**：宿主白名单写死 `WINWORD.EXE` / `EXCEL.EXE` / `POWERPNT.EXE`，其余进程（WPS 的
  `wps.exe`/`et.exe`/`wpp.exe` 等）在 `OnConnection` 里直接返回 —— 不建定时器、不建窗口、不介入。
  WPS 自带图片放大预览，不需要本插件（详见第二节末尾）
- 安装需要**管理员权限**（COM 注册必须写 `HKLM`，原因见第六节"踩过的坑"），`uninstall.bat` 可一键清干净
- 不依赖 VSTO 运行时、不需要 Office SDK / Visual Studio；只引用 Office 自带的
  `extensibility.dll`（`IDTExtensibility2` 的官方 PIA，随 Office 安装，无需另外部署）

---

## 一、目录结构

```
imagezoom-addin/
├─ src/ImageZoomAddin.cs    源码（单文件，约 4400 行，C# 5 语法，.NET Framework 4.x）
├─ build.bat                用 Windows 自带的 csc 编译出 ImageZoomAddin.dll
├─ install.bat / install.ps1    注册 COM 类 + 挂 Word/Excel/PowerPoint 加载项（自动请求管理员权限）
├─ uninstall.bat / uninstall.ps1 卸载
├─ check.bat / check.ps1    自检：DLL 版本、注册表、COM 激活、故障记录、日志一次看完
├─ reset.bat / reset.ps1    「要禁用才能打开」时的急救：清除 Office 故障记录 + 复位 LoadBehavior
├─ installer/officeimage.iss     Inno Setup 安装脚本（打包成给最终用户的单文件 exe）
├─ installer/使用说明.txt        安装后放进安装目录、给用户看的说明
├─ assets/                  图标与安装向导侧边图（由 tools/makeicon.py、makewizard.py 生成）
├─ dist/officeimage.exe     ★ 交付物：双击即装的单文件安装包
├─ tools/                   诊断脚本（C# 探针 + Python，定位 COM 激活/vtable/坐标问题用，日常不需要）
├─ README.md                本文件
└─ 技术方案.md               路线选型、取图链路、产品化建议
```

---

## 二、安装（三步）

1. 双击 **`build.bat`** —— 编译。成功会打印 `BUILD OK` 和 DLL 路径。
2. 双击 **`install.bat`** —— 会弹 UAC 提权，点"是"。成功后打印 `COM activation: OK` 和 `DONE.`
3. **关闭所有 Word / Excel / PowerPoint 窗口**，再重新打开。

> 安装为什么要管理员：托管 COM 类必须注册在 `HKLM\SOFTWARE\Classes`，
> 写在 `HKCU` 会激活失败（详见第六节）。这一步不做就永远看不到放大镜。

> DLL 所在文件夹**不要移动**。注册信息里记的是绝对路径，搬家后重跑一次 `install.bat` 即可。

### 给最终用户：一个 exe 装完（`dist\officeimage.exe`）

上面三步是**给开发者自己调试用的**。要发给用户，用打包好的安装程序：

| | |
|---|---|
| 文件 | `dist\officeimage.exe`（2.36 MB，单文件，DLL 已经压缩在里面） |
| 系统 | Windows 7 / 8 / 8.1 / 10 / 11，32 位和 64 位都可以（引导程序 PE 子系统 6.1） |
| Office | 2010 及以上，Word / Excel / PowerPoint 任一或全部；32 位、64 位 Office 都能加载 |
| 前置 | .NET Framework 4.0+（Win8 以上自带；Win7 没装的话安装程序会明确提示） |

用户侧动作就三步：双击 → 点 UAC 的"是" → 一路"下一步"。
安装到 `C://Program Files\OfficeImage`，**卸载项自动出现在"程序和功能"**：

```
设置 → 应用 → 已安装的应用 → "OfficeImage 图片放大预览（Word / Excel / PowerPoint）"
Windows 7：控制面板 → 程序和功能
```

想静默批量部署（域/脚本分发）就加参数：

```bat
officeimage.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
```

**它到底写了什么**（和 `install.ps1` 等价，只是换成标准安装程序）：

1. `ImageZoomAddin.dll` 复制到 `{app}`，并把 `CodeBase` 指到这个最终路径
2. COM 注册写 **HKLM64 和 HKLM32 两个视图** —— 32 位 Office 会把
   `HKLM\Software\Microsoft\Office` 重定向到 Wow6432Node，少写一个视图就有一半机器加载不到
3. `Word / Excel / PowerPoint` 三个 `Addins\ImageZoomAddin.Connect` 都写 `LoadBehavior=3`
   （同样两个视图各一套，共 6 处）
4. 清掉 Office 里 `Resiliency\DisabledItems / StartupItems / DisabledAddins` 中带本插件名字的
   二进制记录 —— 否则加载项被 Office 拉黑过一次就再也起不来
5. 写 `Uninstall` 键（显示名、版本、发布者、图标、大小、卸载命令），装的时候顺手记下改了哪些键，
   卸载时原样还原

**重新打包**（改了代码之后）：

```bat
python tools\makeicon.py                              :: 只在换图标时需要
python tools\makewizard.py                            :: 只在换向导侧边图时需要
"%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" installer\officeimage.iss
```

版本号不用改脚本：`officeimage.iss` 里用 `GetVersionNumbersString("..\ImageZoomAddin.dll")`
读 DLL 的版本，`AppVersion`、`InprocServer32\Assembly` 全都跟着走。
（前提是源码里 `AssemblyVersion` 和 `AssemblyFileVersion` 保持一致，目前都是 `1.0.0.0`。）

> **Inno Setup 是第三方工具**，本机已装在 `C://Program Files (x86)\Inno Setup 6\`。
> 6.7 的免费版编译时横幅写的是 `Non-commercial use only`：自己用、免费分发没问题，
> 如果这个插件要**随商业产品分发或直接卖**，按它的授权需要买一份商业许可。
> 不想用 Inno 的话，`installer/officeimage.iss` 里的注册表清单可以直接搬去 NSIS/WiX 重做。

> 覆盖安装（升级）也走同一个 exe：检测到已安装会直接覆盖，不会产生第二个"已安装程序"条目。
> Office 正在运行时，安装程序会提示先关闭 Word/Excel/PowerPoint（DLL 被占用）。

### 关于 WPS：本插件不会在 WPS 里生效

WPS 的文字/表格自带"图片放大预览"，不需要本插件。为了确保它绝不会在 WPS 里冒出来，做了两层隔离：

1. **注册层**：`install.bat` 写 `HKCU\Software\Microsoft\Office\{Word|Excel|PowerPoint}\Addins`，
   安装包 `officeimage.exe` 改成写 `HKLM` 同名位置（全机器生效），两者都不会碰 WPS 的键；
   同时会主动清理 `HKCU/HKLM\Software\Kingsoft\Office\WPS\Addins` 下同名的残留项（正常情况本来就没有）。
2. **代码层（关键）**：`OnConnection` 第一件事就是查宿主进程名，
   只有 `WINWORD.EXE` / `EXCEL.EXE` / `POWERPNT.EXE` 放行，其余一律直接 `return`：

   ```csharp
   string host;
   if (!IsSupportedHost(out host))   // 只认 WINWORD / EXCEL / POWERPNT
   {
       Log.Write("OnConnection: 跳过（宿主不受支持：" + CurrentProcessName() + "）");
       return;                       // 不建 Timer、不建窗口，等于完全不介入
   }
   ```

   为什么不能靠 `Application.Name` 判断：部分 WPS 版本会复用上面那份
   `Microsoft\Office\*\Addins` 注册把加载项拉起来，而它暴露的 `Application.Name`
   五花八门（"WPS Writer"/"WPS 表格"…），名字匹配容易漏。进程名最干脆。

验证过四种进程名的判定结果（`tools/hostprobe.exe` 会反射调用 `IsSupportedHost`）：

| 进程名 | 结果 |
|---|---|
| `WINWORD` | 放行，`host=Word` |
| `EXCEL` | 放行，`host=Excel` |
| `POWERPNT` | 放行，`host=PowerPoint` |
| `wps` / `et` / `wpp` | **拒绝** |
| 其他 | **拒绝** |

---

## 三、使用

| 操作 | 效果 |
|---|---|
| **单击一张图片**（文档 / 工作表 / 幻灯片） | 图片**右下角**浮出圆角方框按钮组（不影响你的选中状态） |
| **点击放大镜** | 弹出预览窗，默认按**原图尺寸**（1 图片像素 = 1 屏幕像素）显示 |
| **点击放大镜下方的剪刀图标** | 在图标右侧弹出二级菜单（WPS 样式）：按形状裁剪 / 按比例裁剪 / 自由裁剪 / 重设形状和大小（三个宿主都有） |
| **再下方的弧形箭头图标** | 弹出旋转菜单：向右旋转 90° / 向左旋转 90° / 水平翻转 / 垂直翻转（三个宿主都有） |
| **滚轮** | 以光标位置为中心缩放（0.02x ~ 40x） |
| **按住左键拖动** | 平移图片，看任意位置的细节 |
| **双击** | 在「1:1 原图」和「自适应窗口」之间切换 |
| **Esc** | 关闭预览窗 |
| `0` / `1` / `+` / `-` | 自适应 / 1:1 / 放大 / 缩小 |
| `Ctrl+T` | 预览窗置顶开关 |

按钮的排布：放大镜贴住图片右下角，裁剪在它正下方，旋转在最下面，竖排三个。默认落在图片**右侧外面**（图片内侧的右下角在裁剪状态下正好是拖拽手柄，压上去会挡住操作）；图片紧贴屏幕右边时自动缩回图片内侧。按钮位置每 200ms 跟随图片的屏幕矩形刷新 —— **缩放、拖动图片、滚动页面/工作表、切换幻灯片，按钮都跟着走**。

按钮外观照 WPS 的图片快捷按钮做：30×30 的**圆角方框**（圆角半径 6，约为边长的 20%），白底 + 1px 浅灰描边（`#DEE2EA`）。

鼠标悬停时：底色转成浅蓝 `#E0EBFB`、描边转 `#7498D6`，图标也整体压深一档（深灰 `#3A3A3A` → `#101010`，品牌蓝 `#346ED4` → `#184CAA`），让"鼠标确实停在这儿"这件事一眼看得出来。

图标是手绘的线条图标（深灰打底 + 品牌蓝点缀），全部落在 **x.5 的半像素网格**上、线宽正好 1px：

- **裁剪**：两条 L 形角标组成的 ⌗ 符号 + 一道蓝色斜线；
- **放大镜**：圆镜片 + 里面一个加号 + 右下短手柄；
- **旋转**：小方框 + 蓝色"从右侧上去 → 拐弯 → 向左"的箭头。

### 按钮是怎么画"精细"的

这是这块最容易被忽略的地方，做法在 `MagnifierButton`：

1. **用分层窗口（`WS_EX_LAYERED` + `UpdateLayeredWindow`）而不是 `Region` 裁剪。**
   WinForms 的 `Region` 是 1 位掩码 —— 圆角只能一格一格地切，放大看就是阶梯锯齿，
   这也是"外框边缘毛糙、像分辨率很低"的真正原因（系统 DPI 是 96，不存在 DPI 拉伸问题）。
   分层窗口的内容是一张带 alpha 的位图，圆角边缘可以有真正的半透明过渡。
2. **8 倍超采样再缩小。** 先在 240×240 上按 30×30 的逻辑坐标作图，
   再按 8×8 像素做**盒式平均**缩回 30×30。不用 GDI+ 的插值：它缩小 8 倍会漏采样，
   1px 的线缩回来是半透明的一团灰；盒式平均是整数倍缩小的正解，对齐好的 1px 线能 100% 实心。
3. **图标坐标对到半像素**（0.5 / 2.5 / 13.5…）且线宽正好 1.0，
   这样 1 个像素被完整覆盖，缩回来是实心深灰而不是发虚的灰边。

### 裁剪二级菜单（Word / Excel / PowerPoint）

点剪刀图标，右侧弹出面板（再点一次收起，点面板外面也收起）：

| 位置 | 内容 |
|---|---|
| 标签页一「按形状裁剪」 | 矩形组 6 个 + 基本形状 28 个（椭圆/菱形/六边形/五角星/箭头/心形/云/月牙…） |
| 标签页二「按比例裁剪」 | 1:1、4:3、3:4、16:9、9:16、3:2、2:3、5:4、4:5、21:9、原图比例 |
| 底部第一行「自由裁剪」 | 进宿主原生裁剪状态，可拖四边四角 |
| 底部最后一行「重设形状和大小」 | 形状还原成矩形、裁剪清零、尺寸回到默认值 |

实现方式（每条都验过）：

- **按形状裁剪**，三个宿主三条路：
  - **Word**：改图片 drawing 的 `prstGeom`（预设几何）再 `InsertXML` 写回 —— 这是唯一既能真裁成形状、又保持**嵌入型**不变的办法（Word 里 `Shape.AutoShapeType` 对图片无效，赋值被静默丢弃，读回永远是 `-2`）。
  - **PowerPoint**：直接 `Shape.AutoShapeType = msoShapeXxx`。PowerPoint 的 `p:pic` 里 `a:prstGeom` 本来就是可写的，设完 `Type` 仍是 13（图片）、`PictureFormat` 照样可用，所以「先裁成椭圆、再改 16:9」可以叠加（实测 9:16 出来正好 0.5625）。
  - **Excel**：没有这条管道，见下面「Excel 的形状裁剪」。
- **按比例裁剪**：只改 `PictureFormat.CropLeft/Right/Top/Bottom`，不动形状、不换对象，随时可以改回来。注意它的单位是「原图 100% 尺寸的磅值」而不是当前尺寸，代码里用一次探针量出基准再迭代收敛（实测误差 < 0.2%）。
- **重设形状和大小**：`ExecuteMso("PictureResetAndSize")`，宿主自带的「重设图片和大小」，形状、裁剪、尺寸一起还原。

#### Excel 的形状裁剪

Excel 的图片（`x:pic`）没有任何可写的几何接口，也没有"把图片裁剪成形状"的命令，所以走的是等价实现：

1. 按原图的 `Left/Top/Width/Height` 新建一个自选图形（`msoAutoShapeType` 编号由 `tools/shapemap.cs` 实测反查，见 `tools/shapemap.txt`）；
2. `Shape.Fill.UserPicture(临时图片文件)` 把图片填进去、去掉边框 —— 实测 `Fill.Type` 变成 6（图片填充），图片字节复用 `xl/media` 里那一份，**文件不会因此变大**；
3. 原图 `Visible = msoFalse` 藏起来，形状 `Name = 原名 + "#shape"`、`AlternativeText = IZADDIN:SHAPE:<prst>|<原名>` 做标记 —— 有了这个标记，插件仍然把这块图形当作"图片"来显示放大镜/剪刀，并且能随时把原图找回来；
4. 「重设 / 换比例 / 自由裁剪」会先把图形还原成原图（删掉图形、原图恢复可见），再执行对应动作；换形状则直接换成新的图形。



放大到 100% 以上时使用最近邻插值（像素清晰不糊），缩小到 100% 以下使用高质量插值（不出现锯齿）。

### 旋转 / 翻转（Word / Excel / PowerPoint）

点第三个按钮（弧形箭头 ↻）弹出四项，内容与 Office「图片格式 → 排列 → 旋转」下拉一致：

| 菜单项 | 效果 |
|---|---|
| 向右旋转 90° | 顺时针 90°，可反复点（4 次回到原样） |
| 向左旋转 90° | 逆时针 90° |
| 水平翻转 | 左右镜像 |
| 垂直翻转 | 上下镜像 |

两个宿主的实现完全不同：

- **Excel**：`Shape.Rotation`（顺时针角度，累加后取模）+ `Shape.Flip(0 / 1)`（msoFlipHorizontal / msoFlipVertical），
  和 Office 自己用的是同一套对象模型属性。裁剪成形状得到的「图片填充自选图形」也是 Shape，一样能转。
  旋转过之后 `Left/Top/Width/Height` 仍是"没转之前"的框，屏幕上的外接矩形会变，
  所以 `GetExcelShapeRect` 里按 `rot` 折算了外接矩形（见踩坑 14），按钮才不会贴着旧框跑。
- **PowerPoint**：和 Excel 一样用 `Shape.Rotation` + `Shape.Flip(0 / 1)` —— 三个宿主里唯一现成好用的。
  有个小特点：PowerPoint 会把「旋转 + 翻转」自己归一化，90° 的图 `Flip(0)` 之后 `Rotation` 读回来是 270
  （视觉上是同一个结果）。所以每次都"读当前值再累加"，绝不缓存角度状态。
- **Word**：图片对象**没有**可用的旋转 API（`InlineShape` 没有 `Rotation`；`Shape.Rotation` 只管浮动图，
  也没有 Flip）。Word 自己存旋转的方式就是在 `pic:spPr/a:xfrm` 上写 `rot` / `flipH` / `flipV`
  （单位 1/60000 度，一周 21600000），所以这里走和「裁剪为形状」同一条管线：
  取 `Range.WordOpenXML` → 改属性 → `InsertXML` 写回。嵌入型图片改完仍是嵌入型。

---

## 四、原图是怎么取到的

这是本插件的技术核心：**不读"显示出来的图"，而是读文件里存的原始图片字节。**

| 优先级 | 场景 | 来源 | 是否原图 |
|---|---|---|---|
| 1 | Word 嵌入式图片 | `Range.WordOpenXML` → Flat OPC 包里的 `image/*` 部件（base64） | ✅ 原图字节 |
| 2 | Excel 任意图片 | `Workbook.SaveCopyAs` 临时 xlsx → 解包 → `xl/drawings/*.xml` 按 `Shape.Name` 命中 → `r:embed` → `xl/media/*` | ✅ 原图字节 |
| 2.5 | PowerPoint 任意图片 | `Presentation.SaveCopyAs` 临时 pptx → 解包 → `ppt/slides/*.xml` 按 `p:cNvPr/@id` 命中（`cNvPr/@id` 恒等于 `Shape.Id`）→ `r:embed` → `ppt/media/*` | ✅ 原图字节 |
| 3 | 兜底 | 复制图片 → 剪贴板里的 `PNG`/`JFIF` 原始格式流 | ✅ 通常也是原图 |
| 4 | 最后兜底 | 剪贴板渲染位图 | ⚠️ 可能非原图（标题会标注"渲染位图"） |

预览窗标题会显示 `原图 1920×1080px · 文档中显示 240×135px · 来源：OOXML 原始字节`，一眼能看出是不是真原图。

---

## 五、已知限制

1. **Word 浮动（环绕）图片**：没有公开 API 能直接定位它的图片部件，走的是剪贴板链路。
   绝大多数情况仍是原图字节；若标题显示"渲染位图"，说明这张图只能拿到渲染分辨率。
2. **链接到外部文件的图片**、**OLE 对象**、**Excel 365 的"放置在单元格中"图片**（rich value）不在支持范围。
3. **Excel / PowerPoint 首次预览会慢一点**：为了取原图要把工作簿/演示文稿另存一份临时副本再解包，
   大文件会慢一些（临时文件用完即删）。PPT 通常比 xlsx 大，感知更明显。
4. 取原图兜底路径会操作一次剪贴板（会覆盖你剪贴板里的内容），仅在前面几条链路都失败时发生。
5. 预览窗是独立顶层窗口，高 DPI（150%/200%）缩放下可能略糊（WinForms 未做 per-monitor DPI 适配）。
6. **Excel 的「裁剪为形状」换来的是自选图形**：它在 Excel 里已经不是"图片"对象了（类型 1 + 图片填充）。插件靠 `IZADDIN:SHAPE:` 标记继续把它当图片用，但 Excel 自己的「图片格式」相关命令对它无效；把它复制到别的工作簿后，标记里的原图可能找不到，"重设"会提示找不到原图。
7. **Excel 重设尺寸的基准会被"另存一份临时副本"改掉**：实测同一张 800×600px 的图，插入后设成 220×160，只要执行过一次取原图（预览或裁剪为形状都会触发 `SaveCopyAs`），Excel 就把 220×160 当成了 100%，之后「重设」回到的是 220×160。用户看到的仍是"回到裁剪之前的样子"，只是不会放大回原始像素尺寸。
8. **PowerPoint 里选中"组合"时没有按钮**：组合被选中时 `Selection.Type` 是 0（`ppSelectionNone`），
   `ShapeRange` 一并不可用，插件无法判断你想操作组里哪一张。要改组合里的图，先取消组合或双击进组。
9. **PowerPoint 只在普通视图 / 幻灯片视图下工作**：幻灯片浏览、备注、放映视图里不显示按钮（放映时
   `ActiveWindow` 都不可用）。单选多张图片时按第一张定位（与 Excel 行为一致）。
10. **PowerPoint 的占位符图片**靠 `PlaceholderFormat.ContainedType == 13` 识别；如果是给占位符
    做了"图片填充"（而不是插入图片），`p:sp` 的 blip 仍能取到，但少数版式会退到剪贴板兜底。
11. 图片是超大图（如 8000×6000）时解码占内存约 200MB。
12. 「裁剪为形状」是改 OOXML 实现的：极少数非标准 drawing（比如自定义几何 `custGeom`）
    找不到 `prstGeom`，会弹提示让你改用自带裁剪，不会破坏文档。Word 里处理浮动图片时，
    内部会先把它临时转成嵌入型再转回浮动，环绕方式能还原，但锚点位置可能微调 1~2pt；
    嵌入型图片（绝大多数情况）没有这个问题。
13. **按钮定位靠宿主自己的坐标换算**：Excel 走「屏幕 DPI / 72 × 工作表缩放」推算（见踩坑 13），
    PowerPoint 直接问 `PointsToScreenPixels`（见踩坑 18）。96 DPI 下各档缩放都实测正确；
    系统缩放 150%/200% 的情况没有实机验证过，若按钮有偏移，把
    `%LOCALAPPDATA%\ImageZoomAddin\addin.log` 里的 `magnifier:` 行发我。
14. **Excel 拖动/缩放图片的过程中按钮会滞后**：那几秒 Excel 拒绝 COM 调用（`0x800AC472`），
    拿不到新坐标，只能松手后的下一轮轮询（≤200ms）才归位。

---

## 六、排障

### 先跑 `check.bat`

它会把「DLL 版本 / 注册表内容 / COM 激活测试 / 日志 / Office 进程」一次性打出来，
哪一步红了就是哪一步的问题：

| check.bat 显示 | 含义 | 处理 |
|---|---|---|
| `HKLM 64-bit: MISSING` | COM 类没注册到 HKLM（最常见） | 以管理员身份重跑 `install.bat` |
| `HKCU registration also present` | 旧版遗留的 HKCU 注册 | 无害，但真正生效的是 HKLM |
| `COM activation test: FAILED` + `0x80070002` | mscoree 找不到程序集 | 几乎一定是"只注册在 HKCU"，重跑 `install.bat`（管理员） |
| `LoadBehavior = 2` | 曾经加载失败，Office 把它降级禁用了 | 重跑 `install.bat`（会写回 3），再重启 Office |
| `Office resiliency records` 有红字 | Office 记了故障项，会**跳过**该加载项 | 重跑 `install.bat`（会自动清除），再重启 Office |
| `NO LOG FILE` | COM 对象从未被创建 = 卡在激活阶段 | 不是代码逻辑问题，看上面两条 |
| 日志有 `ctor:` 但没有 `OnConnection` | **没引用官方 PIA**，CLR 封送 `custom` 参数失败（见坑 4） | 确认 `using Extensibility;` 且 build.bat 引用了 extensibility.dll |
| 有 `OnConnection` 但没有 `diag#` | 加载项进来了，但 200ms 定时轮询没跑起来 | 把日志发我，需要换触发机制 |
| 有 `diag# ... Selection.Type=1` | 轮询正常，但当前选中的不是图片 | Word 应为 7（内嵌图）/ 8（浮动图）；PowerPoint 应为 2（`ppSelectionShapes`） |

### 踩过的坑（已修，共十九处）

**坑 1：程序集版本号不匹配**

**症状**：装了之后 Word/Excel/PowerPoint 都没反应，`LoadBehavior` 变成 `2`。

**根因**：`csc` 在没有 `AssemblyVersion` 特性时编译出的程序集版本是 **`0.0.0.0`**，
而 `install.bat` 早期硬编码写成 `Version=1.0.0.0`。mscoree 按错误版本号找不到程序集，
COM 激活失败（`0x80070002`）。

**修法**：源码显式声明 `[assembly: AssemblyVersion("1.0.0.0")]`；安装脚本改成从 DLL 读真实版本。

**坑 2（真正的元凶）：托管 COM 类注册在 HKCU 里永远不会被激活**

**症状**：修好版本号后，`0x80070002` 依旧，`OnConnection` 一次都没跑过。

**定位过程**（`tools/` 下留有脚本）：

1. 解析 DLL 元数据确认 `AssemblyVersion=1.0.0.0`，与注册表声明**完全一致** → 版本号不是原因
2. 用 7 种不同写法注册（RegAsm 标准布局、版本子键、纯部分名、短路径、去 CodeBase 走 GAC）
   → **全部同一个 `0x80070002`**
3. 拿 `System.dll`（GAC 里的系统程序集）做对照 → **也失败**
4. 扫描系统里已有的托管 COM 类 → 发现 HKLM 下的 `mscorlib 2.0` 类**激活成功**
5. 把**完全相同的配置**写进 HKCU → **失败**
6. 把我们的类从 HKCU 挪到 HKLM → **成功（`0x00000000`）**

**结论**：`mscoree` 在把 CLSID 解析成程序集时只认 **HKLM**。
`HKCU\Software\Classes\CLSID` 下的托管 COM 类，`CoCreateInstance` 一律返回
`0x80070002 系统找不到指定的文件`——而且**不报任何其它错**，Office 只是静默不加载。

**修法**：
1. `install.ps1` 把 CLSID 与 ProgID 写入 `HKLM\SOFTWARE\Classes`（同时写 `Wow6432Node` 兼容 32 位 Office）
2. `install.bat` 用 `Start-Process -Verb RunAs` 自动请求管理员权限
3. 额外清除 `Resiliency\StartupItems` 里的故障记录——**这一条同样致命**：
   Word 一旦把某个加载项记进 StartupItems，之后即使 `LoadBehavior=3` 也会跳过它

**验证结果**：`QI IUnknown / IDispatch / IDTExtensibility2` 全部 `S_OK`，
`CoCreateInstance` 返回 `0x00000000`，日志出现 `ctor: add-in object created`。

**坑 3（让 Excel 弹"禁用才能打开"的那一个）：Office 调到的是 `ToString`，宿主直接崩**

**症状**：
- 打开 Excel 弹窗「此加载项导致严重问题，是否禁用它？」，点了禁用才能进 Excel
- `Resiliency\DisabledItems` 里出现 `...imagezoomaddin.dll!imagezoomaddin.connect`
- 日志只有 `ctor: add-in object created`，**永远没有 `OnConnection`**
- Windows 事件日志里 Word 和 Excel 崩在**同一个** clr.dll 偏移，退出码 `80131506`

**根因**：`IDTExtensibility2` 在 `MSADDNDR.TLB` 里是 **dual 接口**
（`interface IDTExtensibility2 : IDispatch`），vtable 必须是：

```
slot 0-2  IUnknown     QueryInterface / AddRef / Release
slot 3-6  IDispatch    GetTypeInfoCount / GetTypeInfo / GetIDsOfNames / Invoke
slot 7-11 五个方法      OnConnection / OnDisconnection / OnAddInsUpdate / OnStartupComplete / OnBeginShutdown
```

Office 按 vtable 直调 **slot 7**。而 `[ClassInterface(ClassInterfaceType.AutoDual)]`
会让类多出一个 `_Connect` 类接口，`tools/typedump.py` 实测它的布局是：

```
slot 7  ToString      slot 8  Equals     slot 9  GetHashCode    slot 10 GetType
slot 11 OnConnection  slot 12 OnDisconnection ...
```

**`System.Object` 的四个公共方法被排在了最前面，`OnConnection` 被挤到 slot 11。**
Office 调 slot 7 拿到的是 `ToString`（无参、返回 BSTR），签名完全对不上 →
CLR 致命错误 `0x80131506` → 宿主进程崩溃。
（`dotnet/runtime#38950` 记录的正是 "instead of calling OnConnection(),
it is calling the object's ToString()"。）

如果不用 `AutoDual`，默认的 `AutoDispatch` 会生成一个只有 IDispatch 7 个槽位的类接口，
slot 7 落到堆内存上 —— 一样崩。**两条路都错，因为问题在于「让类接口存在」本身。**

**修法**：类不要有类接口，任何接口指针的 slot 7 都必须是 `OnConnection`。

```csharp
using Extensibility;                                 // 用官方 PIA 的接口定义

[ClassInterface(ClassInterfaceType.None)]            // 关键：不生成类接口
public class Connect : IDTExtensibility2 { ... }
```

> 注意：`ClassInterfaceType.None` **只解决 vtable 布局**，它不解决参数封送。
> 自己手写 `IDTExtensibility2` 时，即使 `verify.py` 显示 slot 7 能跑通，
> Office 走 IDispatch 调用时仍会在封送 `custom` 参数时崩 —— 那就是坑 4。
> **两个都要修：`None` + 官方 PIA 接口。**

**验证**：最可靠的方式是用真实宿主跑一遍（见坑 4 末尾的 PowerShell 片段）。
`tools/` 下的 `typedump.py` / `verify.py` 仍有参考价值（能看到 `_Connect` 把
`ToString` 摆在 slot 7），但它们手工拼 `custom` 参数，会误报 `E_INVALIDARG`，
不要拿它们的结论当最终判据。

**坑 4（真正让宿主每次都崩的那一个）：自己手写 `IDTExtensibility2` 接口，CLR 封送不出正确的 `custom` 参数**

**症状**：日志每次都有 `ctor:`，**永远没有 `OnConnection: enter`**；Word/Excel 弹
「此加载项出现问题，是否禁用」，Windows 事件日志里是
`错误模块 clr.dll` + `退出代码 80131506`（COR_E_EXECUTIONENGINE），而且**每次崩在同一个 clr.dll 偏移**。

**关键判据**：崩点固定在 clr.dll 的同一偏移 → 崩在**参数封送阶段**，不是 vtable 布局，
也不是方法体里的代码（方法体第一行就是 `Log.Write("OnConnection: enter")`，它没执行）。

**根因**：`IDTExtensibility2` 的 `custom` 参数在 IDL 里是
`[in,out] SAFEARRAY(VARIANT) *custom`。**自己手写这个接口（哪怕写了 `ref Array`、
`InterfaceIsDual`、`AutoDual`）都不够**：CLR 需要接口类型上的 `ITypeInfo` 才能正确封送，
只有官方 PIA 里的定义带 `[ImportedFromTypeLib] + [TypeLibType(0x1040)]`，才会让 CLR
改用类型库规则封送。

**修法（最终方案）：不要自己定义接口，直接引用微软官方 PIA。**

```csharp
using Extensibility;   // 官方 PIA：IDTExtensibility2 / ext_ConnectMode / ext_DisconnectMode
```

`build.bat` 里加上引用：

```
/r:"%WINDIR%\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll"
```

> 这个 PIA 随 Office / Visual Studio 一起安装，Windows 机器基本都有。
> 走过弯路：`[MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]`
> 编译能过，但运行时元数据里查不到 `VarEnum`，无效。

改完后用真实的宿主验证（最可靠，不依赖任何手工拼参数的脚本）：

```powershell
$w = New-Object -ComObject Word.Application   # 插件随 Word 启动自动加载
Start-Sleep 4
Get-Content "$env:LOCALAPPDATA\ImageZoomAddin\addin.log"
```

日志应出现 `ctor:` → `OnConnection: enter` → `StartPolling: timer started` → `OnConnection: done`。

**坑 5：`Selection.InlineShape`（单数）根本不存在**

**症状**：加载项一切正常，选中图片后放大镜不出现，日志刷
`RuntimeBinderException: "System.__ComObject" 未包含 "InlineShape" 的定义`。

**根因**：Word 的 `Selection` 对象只有 `InlineShapes`（集合），没有单数的 `InlineShape`。
内嵌图形必须从集合里取：

```csharp
dynamic inlines = sel.InlineShapes;      // ✅
dynamic shape   = inlines.Item(1);
// dynamic shape = sel.InlineShape;      // ❌ 运行时必抛异常
```

**坑 6：`wdSelectionInlineShape` 是 **7**，不是 3**

**症状**：同上，选中图片后放大镜不出现，日志里 `Selection.Type=7` 但代码在等 3。

**根因**：Word 的 `WdSelectionType` 枚举里 `wdSelectionInlineShape = 7`，
`3` 是 `wdSelectionFrame`（图文框）。之前照抄了一个错误常量，导致分支永远进不去。

| 常量 | 值 | 含义 |
|---|---|---|
| `wdSelectionIP` | 1 | 插入点 |
| `wdSelectionNormal` | 2 | 选中了普通文本 |
| `wdSelectionInlineShape` | **7** | 选中了嵌入式图片 |
| `wdSelectionShape` | **8** | 选中了浮动图片/形状 |

**坑 7：把 DLL 编到临时文件名，程序集会改名，Office 直接加载失败（`0x80131040`）**

**症状**：Office 不再崩溃，但加载项完全不生效，日志里连 `ctor` 都没有，`LoadBehavior` 从 3 掉回 2。
本地跑 `python tools/dispinvoke.py byref-array` 会看到 `CoCreateInstance hr=0x80131040`
（`FUSION_E_REF_DEF_MISMATCH`：程序集引用与定义不匹配）。

**根因**：csc 用 `/out` 的**文件名**决定程序集的简单名。为了绕开"文件被占用"改成
`/out:ImageZoomAddin.new.dll` 编译，程序集名就变成了 `ImageZoomAddin.new`，
而注册表里 `Assembly=ImageZoomAddin`，mscoree 找不到相符的程序集，激活直接失败。

**修法**：临时输出必须放在**同名子目录**里，文件名保持 `ImageZoomAddin.dll`：

```
.obj\ImageZoomAddin.dll      ← 编译到这里
ImageZoomAddin.dll           ← 成功后替换过去
```

**坑 8：Office/WPS 开着的时候编译，报「另一个程序正在使用此文件」**

**根因**：宿主进程把已加载的 DLL 内存映射了，不能覆盖写。这类文件通常**还能改名**
（映射句柄允许 delete-share），但绝对不能再写。

**修法**：`build.bat` 现在先编到 `.obj\`，再 `move` 覆盖；失败时把旧文件改名成
`ImageZoomAddin.previous.dll` 让位，然后把新文件放进去，最后尝试删掉旧文件。
旧文件如果还被占用删不掉，留着也无害，下次编译会自动清理。

> 顺带一个坑：加载失败一次，Word 会把 `LoadBehavior` 降成 2 并**不再尝试加载**。
> 修好 DLL 后要用 `reset.bat`（或手动把 `LoadBehavior` 改成 3）才会重新加载。
> Excel 是独立的一份注册，可能仍然是 3，两个宿主分别看。

**坑 9：Word 里没有「裁剪成形状」的 API（三处都试过）**

做 WPS 样式的裁剪菜单时踩到的，结论都来自实测：

1. `Shape.AutoShapeType = msoShapeOval` —— **对图片形状无效**。
   `InlineShape.ConvertToShape()` 出来的是 `msoPicture`（Type=13），
   `AutoShapeType` 读回来永远是 `-2`（msoShapeMixed），赋任何值都被**静默丢弃**，不报错。
2. `Shapes.AddShape` + `Fill.UserPicture` —— 能画出"图片填充的形状"，但那是**新建的浮动形状**，
   原来的嵌入型图片得删掉，正文会重排，而且要自己保存原始字节才能重做，代价太大。
3. **可行的是改 OOXML**：取 `Range.WordOpenXML`，把第一个
   `<a:prstGeom prst="rect">` 换成目标形状名，再 `Range.InsertXML()` 写回去。
   实测图片仍是嵌入型（`InlineShapes.Count=1`、`Type=3`），形状真实生效，
   而且 `PictureResetAndSize` 能一键还原成 `rect`。

> **副作用（必读）**：`InsertXML` 会换掉整个 Range，**选中状态丢失**
> （`Selection.Type` 从 7 变回 1），放大镜和裁剪按钮会跟着消失。
> 写完必须按 `Range.Start` 把新图片找回来重新 `Select()` —— `AfterInsert()` 干的就是这件事。
> 另外浮动图片的 `Shape.Range.WordOpenXML` **返回空串**，所以浮动图要先
> `ConvertToInlineShape()` 改完再 `ConvertToShape()` 转回去。

**坑 10：`PictureFormat.CropLeft` 的单位不是当前尺寸**

想按 16:9 裁剪时按"当前宽度的百分比"算裁剪值，结果裁得完全不对。
实测（3840×2160 的图，文档里显示 320×180pt，`PictureResetAndSize` 后是 576×324pt）：

| 设置 | 结果 | 说明 |
|---|---|---|
| `CropLeft = 57.6` | 可见宽度 320 → 288 | 57.6 / **576** = 10%，掉的是当前宽度的 10% |
| `CropLeft = 41.53` | 可见宽度 320 → 288.6 | 41.53 / 576 = 7.2% |

即：**裁剪值的单位是"原图 100% 尺寸的磅值"**，可见尺寸按比例同步缩放。
而"原图 100% 尺寸"跟插入尺寸、跟图片 DPI 都不是同一个数（这张图插入是 415.3×233.6，
重设回 576×324），不能靠 DPI 换算硬算。

**修法**：用一次探针把它量出来 —— 先裁 10%，看宽度掉了多少，
`W0 = 探针值 / (1 - 裁后宽/原宽)`，再按 `ΔC = (当前宽 - 目标宽) * W0 / 全宽` 迭代两次收敛。
实测 1:1 得 1.0008、16:9 得 1.7769、9:16 得 0.5624（目标 0.5625）。

**坑 11：Excel 里缩放图片后，放大镜停在原地不动**

按钮重排的判断条件原来只有「还是同一张图（Key）+ 两个按钮都可见」。
Word 里拖动手柄时选中对象会被打断（`Selection` 变成别的东西 → 先隐藏 → 松手后重新出现），
所以看起来"跟随"了；Excel 缩放图片时 `Selection` 始终是那个 Shape，Key 一模一样，
于是一次都不会重排，按钮死守在旧位置。

**修法**：判断条件加上**图片在屏幕上的矩形**（`SameRect`，容差 1px 防像素取整抖动）。
顺带把"滚动页面 / 工作表导致图片移屏"也一起修好了 —— 那种情况 Key 同样不变。

**坑 12：Excel 根本没有"裁剪为形状"，`x:pic` 也没有可写的几何**

Word 那条 `prstGeom` + `InsertXML` 的路子在 Excel 完全不存在（没有 `Range.InsertXML`，
`Shape.AutoShapeType` 对图片同样无效）。唯一的等价物是"自选图形 + 图片填充"：

- `Shapes.AddShape(msoAutoShapeType, …)` 的编号不能靠文档记忆，只能实测：
  `tools/shapemap.cs` 把 1..250 逐个建形状、落盘后解析 `drawing1.xml` 里的 `a:prstGeom prst`，
  反查出 `rect=1 / roundRect=5 / ellipse=9 / star5=92 / star6=147 / cloud=179 …`（完整表在 `tools/shapemap.txt`）。
  注意 Excel 里"十字形"的 prst 叫 `plus`（编号 11），不是 `cross`。
- `Shape.Fill.UserPicture(路径)` 只认**文件路径**，所以原图字节要先落到临时文件；
  扩展名按字节头判断（png/jpg/gif/bmp/tif），避免 System.Drawing 重新编码掉画质。
- 填完之后 `Fill.Type` 变成 6，`Shape.Type` 仍是 1（自选图形）——
  **图片对象就此消失**，所以要把原图 `Visible = msoFalse` 藏起来，
  并在 `AlternativeText` 里写 `IZADDIN:SHAPE:<prst>|<原名>` 做标记：
  既让插件继续把它当图片（放大镜/剪刀照常出现），也让「重设」能把原图找回来。
- `Ooxml.FromExcelPackage` 的 `cNvPr → blip` 查找原本只认 `xdr:pic`，
  遇到 `xdr:sp + a:blipFill` 会退化到"整张 drawing 里第一个 blip"（画布上有多张图就取错），
  已补上 `sp` 这一层。

> **顺带发现的副作用**：`Workbook.SaveCopyAs`（预览取原图用的）会把图片当前显示尺寸
> 提交成新的"100% 基准"。实测 800×600px 的图插入后设成 220×160，SaveCopyAs 之前裁 22
> 只掉 8pt（基准 600），之后裁 22 掉满 22pt（基准变成 220），`PictureResetAndSize`
> 也就只能还原到 220×160。这不是 bug 而是 Excel 自己的记账方式，见「已知限制」第 7 条。

**坑 13：Excel 的 `Window.PointsToScreenPixelsX/Y` 不按屏幕 DPI 换算**

这是「放大镜跑到图片里面」的真正原因。用 `tools/xlrect.cs`（纯品红图片 + 截屏找色块做基准）
实测出来的规律非常反直觉：

| | f(0) 斜率 | 图片真实像素斜率 |
|---|---|---|
| 工作表 100% | 1.000 | **1.333**（= 96 DPI / 72） |
| 工作表 50% | 1.000 | **0.667** |
| 工作表 150% | 1.000 | **2.000** |

也就是说 `f(p)` 永远是 `p + f(0)`，**完全不带 DPI 也不带缩放**；
而图片在屏幕上 1 磅其实占 `屏幕DPI / 72 × Zoom` 个像素。
只有 `f(0)` 本身是可信的 —— 它就是工作表原点（A1 左上角）的**屏幕绝对坐标**，
连滚动、缩放带来的偏移都算准了，可以当锚点。

按原写法（`f(left)`、`f(left+w)` 相减当宽高）算出来的矩形，100% 缩放下只有真实大小的 3/4，
于是按钮只走了 3/4 的距离，看着就像"压在图片里面"。

**修法**：自己乘系数，锚点仍用 `f(0)`：

```csharp
int ox = win.PointsToScreenPixelsX(0);          // 可信
int x1 = ox + (int)Math.Round(left * dpiX / 72.0 * zoom);
int w  = (int)Math.Round(shape.Width * dpiX / 72.0 * zoom);
```

验证（`tools/xlplace.cs`，截图找品红块对账）：插入 → 放大 → 平移 → 滚动 → 改工作表缩放
五种状态，`dx/dy` 都是 1px（就是窗口边框），`dw/dh` 为 0；
按钮贴边误差 `left=1 / bottom=1`。

另外两条经验：
- 窗口还没真正画出来时 `f()` 一律返回 0（启动后一两秒会这样），此时算出来是垃圾坐标，
  代码里加了 `ox==0 && oy==0 && f(100)==0` 的判空，宁可这一轮不显示。
- Excel 拖动/缩放图片时会进入模态循环，COM 调用全部抛 `0x800AC472`（"忙"），
  这期间就别想拿到新坐标了 —— 只能等下一条轮询（200ms）。所以拖动过程中按钮会有滞后，
  松手后立刻归位，这是 Excel 的限制，不是没跟随。


**坑 14：Word 图片没有旋转 API，Excel 旋转后 Left/Top 还是旧框**

想做「旋转」快捷按钮时先试了三条路：

| 试法 | 结果 |
|---|---|
| `InlineShape.Rotation` | 属性不存在，`dynamic` 直接抛 `RuntimeBinderException` |
| `Shape.Rotation`（浮动图） | 能读能写，但嵌入型（绝大多数图片）没有 Shape 对象 |
| `Shape.Flip` | Word 的 Shape **没有** Flip 方法（Excel 有） |

Word 自己存旋转的位置其实是 OOXML：`pic:spPr/a:xfrm` 上的 `rot`（1/60000 度）、`flipH`、`flipV`。
所以只能走「裁剪为形状」那条改 XML 的路：`Range.WordOpenXML` → 正则改属性 → `Range.InsertXML` 写回。
改完 `rot="0"` 就不写这个属性，保持和 Word 自己产出的 XML 一致；自闭合的 `<a:xfrm/>` 也要能处理
（`rotprobe.cs` 里 13 条断言全覆盖：累加、回退、翻转可逆、自闭合、多属性只留需要的）。

Excel 那边反过来简单（`Shape.Rotation` + `Shape.Flip` 都有），但有个坑：**旋转后 `Left/Top/Width/Height`
仍然是没有旋转时的框**，屏幕上看到的却是绕中心转过去的外接矩形。90° 时宽高正好对调，
所以按钮会贴着旧框的右下角、看着像"跑到图片外面"。修法是按角度折算外接矩形：

```
bw = |w·cosθ| + |h·sinθ|,  bh = |w·sinθ| + |h·cosθ|   // 中心点不变
```

`tools/xlrotate.cs` 验证：右转→90、再右转→180、左转→90、再左转→0，两个方向翻转都不报错，
裁剪出来的「图片填充自选图形」也照样能转。


**坑 15：`Region` 做圆角，边缘永远是锯齿**

按钮从圆形改成圆角方框后，用户反馈"外框边缘比较毛糙，像分辨率很低"。

先排除了 DPI：系统 `AppliedDPI=96`（100% 缩放），不存在位图被拉伸的问题。
真正的原因是 `Form.Region`：它是**1 位掩码**，只有"这个像素算不算窗口内"两种状态，
圆角只能一格一格地切。30px 的按钮上，半径 6 的圆角就会露出明显的阶梯。

**修法**：给窗口加 `WS_EX_LAYERED`，内容改成 `UpdateLayeredWindow` 提交一张 32bppPArgb 位图。
位图带真正的 alpha，圆角边缘可以是半透明过渡，`Region` 也就不用再设了。
副作用是窗口不再走 `WM_PAINT`，所以 `OnPaint`/`OnPaintBackground` 都留空，
悬停变色时直接重新提交一次位图（`Render()`）。

`tools/layeredprobe.cs` 在真实桌面上验证：`_layeredOk = True`、中心像素纯白、
`(0,15)` 正好是描边色 `#DEE2EA`、四角是背景原色（说明圆角外真的透明）。

**坑 16：超采样缩小用 bicubic，1px 的线会变成半透明灰**

分层窗口之后又发现线条发虚：明明画的是 `#3A3A3A`，看着却是浅灰。

原因有两层：
1. 线宽 0.95 压不满一个像素，任何缩小都会把它摊成两个半透明像素；
2. GDI+ 的 `HighQualityBicubic`/`HighQualityBilinear` 缩小 8 倍时会漏采样，
   还会因为振铃在细线两侧留下淡灰光晕。

**修法**：
- 图标坐标全部对到 **x.5 的半像素网格**、线宽正好 **1.0**，让线完整覆盖 1 个像素；
- 缩小这一步不用 GDI+，自己按 8×8 做**盒式平均**（`Downsample`，`LockBits` 直接算），
  整数倍缩小的正解，对齐好的 1px 线缩回来是 100% 实心。

顺带把圆角半径从 8 降到 6（WPS 的圆角约为边长的 20%，不是四分之一圆），
图标缩放从 1.12 降回 1.0 —— 有 `tools/compare.exe` 和 WPS 截图并排比，这些比例一眼就能看出偏差。


**坑 17：PowerPoint 根本没有 `Application.Selection`**

把 Word/Excel 那套直接搬过来时，`_app.Selection` 一调就抛
`MissingMemberException: ApplicationClass 未包含 Selection 的定义`。

PowerPoint 的选中状态挂在**窗口**上：`Application.ActiveWindow.Selection`。
而且 `Selection.Type` 的取值也和 Word 完全不同 —— 选中形状是 **2**（`ppSelectionShapes`），
不是 Word 那种 7（嵌入图）/ 8（浮动图）；放映时 `ActiveWindow` 直接不可用；
组合被选中时 `Selection.Type` 变成 **0**，`ShapeRange` 也一并不可用。所以这条路必须全程容错。

**坑 18：PowerPoint 的 `PointsToScreenPixels` 斜率和 Excel 正好相反**

Excel 那个坑（见坑 13）是"斜率恒为 1.000，必须自己乘 dpi/72 × 缩放"。
到 PowerPoint 我照着抄，结果按钮偏出去 7 个像素 —— 反过来验才发现：
**PowerPoint 的斜率本身就是 dpi/72 × 视图缩放**，一个系数都不能自己乘。

怎么量出来的：用 `DocumentWindow.RangeFromPoint(x, y)`（给屏幕坐标返回那个位置的 Shape）
从图片中心向四边逐像素外扩，反推出图片真实的屏幕矩形，再跟两种算法比。
600 磅的图在 50% / 150% / 200% 下分别是 400 / 1200 / 1600 px，
`f(L+W) - f(L)` 每次都跟真实值一致，而我自乘出来的值在"适应窗口"下永远差一截。

顺带治好了「适应窗口」的坑：新建演示文稿默认按窗口自适应，`View.Zoom` 只报整数 105%，
内部实际是 104.1%，自己乘 1.4 会算出 840px 而实际只有 833px。交给 `f()` 算就没有这个偏差 ——
实测旋转 0°/90° 两种情况，算出来的矩形与 `RangeFromPoint` 量的**误差都是 0px**。

**坑 19：切到下一页，按钮还赖在上一页**

`GotoSlide(2)` 之后，`Selection.Type` 仍然是 2、`ShapeRange` 还是上一页那张图，
于是按钮继续显示在已经翻走的那页上（约 200ms 后才消失）。

修法：拿到 shape 之后核对 `shape.Parent.SlideID == ActiveWindow.View.Slide.SlideID`，
不一致就当没选中。注意 `DocumentWindow` **没有 `HWND` 属性**（类型库里确认过），
录制宏或探针里想拿窗口句柄只能走 `Process.MainWindowHandle`。


### 其他常见问题

**0. Word/Excel/PowerPoint 弹窗「此加载项出现问题，是否禁用？」/ 必须禁用才能打开**

这是加载项曾让宿主动过崩溃，Office 把它记进了黑名单。恢复步骤：

1. 关掉所有 Word/Excel（任务管理器里确认没有 `WINWORD.EXE` / `EXCEL.EXE`）
2. 双击 **`reset.bat`** —— 清除 `Resiliency\DisabledItems|StartupItems|DisabledAddins` 里
   属于本插件的记录，并把 `LoadBehavior` 复位为 3（只影响本插件，不动其他加载项）
3. 重新打开 Office

> `reset.bat` 只是解除封锁，不修 bug。如果它每次都要跑一遍，说明加载项还在崩，
> 把 `check.bat` 的输出和 `%LOCALAPPDATA%\ImageZoomAddin\addin.log` 发我。

**A. 单击图片没有出现放大镜**
1. `文件 → 选项 → 加载项 → 管理:COM 加载项 → 转到`，确认 `Image Zoom Preview` 已勾选。
2. 若曾经加载失败被 Office 拉黑：`管理:禁用的应用程序加载项 → 转到 → 启用`。
3. 看日志：`%LOCALAPPDATA%\ImageZoomAddin\addin.log`
   - 没有 `OnConnection` 行 → 加载项根本没被加载（注册问题，跑 check.bat）
   - 有 `OnConnection: host=Word` 但之后没有 `diag#` → 定时轮询没跑起来
   - 有 `diag# ... Selection.Type=1` → 轮询正常，但你点的不是图片（Type 应为 7=内嵌图，8=浮动图）

**B. 提示"没能取到这张图片的原始数据"**
说明 4 条链路全失败。日志里会有对应异常，通常是图片为链接/OLE/受保护视图（受保护视图中 Office 会禁用加载项）。

**C. 编译失败**
`build.bat` 窗口里 `error CSxxxx` 那几行复制出来即可定位。

**D. 想确认注册是否成功**
```
reg query "HKCU\Software\Microsoft\Office\Word\Addins\ImageZoomAddin.Connect"
reg query "HKLM\Software\Classes\CLSID\{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"
```
`LoadBehavior` 应为 `0x3`（启动时加载）；CLSID 必须能在 **HKLM** 下查到。

---

## 七、后续可以加的功能（按性价比排序）

1. **Ribbon 按钮**：`实现 IRibbonExtensibility` + 自定义 Ribbon XML，加"放大预览 / 导出原图"按钮和快捷键入口（比浮动按钮更稳，可作为第二条触发路径）。
2. **导出原图**：预览窗加"另存为"，直接落盘 `xl/media` 里的原始文件（做文档瘦身/取素材很实用）。
3. **批量导出文档内所有图片**：一次解包全导出，命名 `image1_1920x1080.png`。
4. **对比模式**：同屏并排显示两张图。
5. **打包成安装程序**：Inno Setup + 代码签名证书，写 `HKLM` 全机安装 + `LoadBehavior=3`，做成可分发的产品。
6. **Per-monitor DPI 适配**：manifest 声明 `PerMonitorV2`，预览窗在 4K 屏下按物理像素 1:1 呈现。

---

## 八、为什么不做成"Office 网页加载项（Office.js）"

如果目标包含 Excel，Office.js 这条路是死的，不是取舍问题：

| 需求 | Office.js 现状 |
|---|---|
| 单击图片后浮出按钮 | Word 只能在任务窗格里做按钮，无法覆盖在文档上；Excel 没有"形状被选中"事件 |
| 原图尺寸预览 | Word `getBase64ImageSrc()` 可取原图，但 Excel 只有 `Shape.getAsImage()`，官方写明是 **96 DPI 渲染结果** —— 图片缩小了就只能拿到小图 |
| 浮动图片 | Word JS API 只暴露 `InlinePicture`，浮动图片不可见 |

所以本插件走 **Windows 桌面版 COM 加载项**。代价是只支持 Windows 桌面 Office（不支持 Mac / 网页版 / 手机版）。
