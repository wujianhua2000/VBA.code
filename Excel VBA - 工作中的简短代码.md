

[TOC]



# 工作中简短代码 EXCEL VBA #



```
Option Explicit

Public Function fmtval(value As Double) As String
    fmtval = Right(Space(8) & Format$(value, "0.000"), 8)
End Function
```

