

[TOC]



# 工作中简短代码 EXCEL VBA #



```
Option Explicit

Public Function fmtval(value As Double) As String
    fmtval = Right(Space(8) & Format$(value, "0.000"), 8)
End Function
```

fmtval 主要应用于 ANSYS 的数据准备，VREAD 需要类似 FORTRAN 一样的格式化后的数据，这个命令可以批量处理。

