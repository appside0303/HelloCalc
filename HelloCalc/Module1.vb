'
' VB.netからcalcを呼び出す例題
'
Module Module1

    Sub Main()
        Dim factory As Object
        Dim loader As Object
        Dim component As Object
        Dim args(0) As Object

        factory = CreateObject("com.sun.star.ServiceManager")
        loader = factory.createInstance("com.sun.star.frame.Desktop")
        component = loader.loadComponentFromURL("private:factory/scalc", "_blank", 0, args)

        Dim sheet As Object
        Dim cell As Object

        sheet = component.getSheets().getByName("Sheet1")
        cell = sheet.getCellByPosition(1, 1)
        cell.setFormula("hello, calc.")
    End Sub

End Module

