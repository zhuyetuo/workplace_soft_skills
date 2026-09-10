演示页（第 9 页）的三个图位，文件名固定如下，重新生成 PPT 即自动嵌入：

  scratch.png    抓挠检出与人工确认   —— 已就绪（从 scratch.mp4 抽帧裁切）
  platform.png   皮肤评估每日跟踪表   —— 已就绪
  tartar.png     口腔牙齿检测         —— 已就绪

图片按原始比例居中嵌入，不会拉伸变形；缺图时降级为「现场演示」面板。
scratch.mp4 是标注审核工作台录屏，现场演示直接播放即可。

生成方式：
  node 生成脚本.js
  python3 postbuild.py "转正述职报告-Toky.pptx" media/poster.png 13,14

postbuild 用来把 pptxgenjs 的灰色播放占位图换成 poster.png（视频真实画面）。
