vbCircleDraw
============

Drawing circle...

随机取点并连接，取中点再连接。
理论论文：http://www.cs.cornell.edu/cv/ResearchPDF/PolygonSmoothingPaper.pdf

随手一写，不保证理论理解没问题= =

速度值越大越慢，实际上是一个空循环，最小值为1.

程序有个bug，退出之前要先点stop，否则进程一直在运行。懒得改了……用的时候自己注意一下就好……