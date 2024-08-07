from spire.presentation import *

# 创建一个 Presentation 对象
pres1 = Presentation()
pres2 = Presentation()


# 加载演示文稿文档
pres1.LoadFromFile(r"D:\PPT_Merge\pythonProject\test备选pptx\西井科技-Q-Tractor v2 20240418.pptx")
pres2.LoadFromFile(r"D:\PPT_Merge\pythonProject\test备选pptx\西井科技介绍 -机场版1211.pptx")

# 打印 pres1 和 pres2 的幻灯片数量
print(f"pres1 initial slide count: {len(pres1.Slides)}")
print(f"pres2 slide count: {len(pres2.Slides)}")

# 遍历第二个演示文稿的幻灯片
for slide in pres2.Slides:
    # 将每张幻灯片添加到第一个演示文稿并保留其原始设计
    pres1.Slides.AppendBySlide(slide)

# 打印合并后的幻灯片数量
print(f"pres1 final slide count: {len(pres1.Slides)}")

# 保存第一个演示文稿
pres1.SaveToFile(r"D:\PPT_Merge\pythonProject\test备选pptx\MergePresentations.pptx", FileFormat.Pptx2019)
pres1.Dispose()
pres2.Dispose()
