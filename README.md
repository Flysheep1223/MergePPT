# MergePPT
Merge a lot of PPTs to one PPT  

程序未响应时请勿关闭，该程序在执行较大的pptx时较慢  

**推荐在使用时打开任务管理器，发现内存（主存）接近爆 100% 立即终止该程序**

该程序主要运行流程：  
STEP1: 从外存中读取指定的pptx副本进入主存  
STEP2: 将副本1作为主pptx，遍历副本2的每一页ppt，按顺序复制粘贴进入副本1   
STEP3: 遍历副本3的每一页ppt，按顺序复制粘贴进入副本1  
... ...  
STEPn: 从主存导出副本1到外存

**用的库自带水印，会在生成的pptx中的每一页ppt左上角打上“Eval
uation Warning : The document was created with Spire.Presenta
tion for Python”，如果需要去除水印需要找源库作者购买**
