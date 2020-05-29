Sub Main()
    Try
        SharedVariable("SALES_ORDER") = RuleArguments("SALESORDER")
        SharedVariable("SALES_ITEM") = RuleArguments("SALESITEM")
        SharedVariable("WORK_ORDER") = RuleArguments("WORKORDER")
        SharedVariable("CADFLOW_ACTION") = RuleArguments("CADFLOW_ACTION")
        SharedVariable("SERVER_INSTANCE") = RuleArguments("INSTANCE")
        SharedVariable("ENVIRONMENT") = RuleArguments("ENVIRONMENT")
        SharedVariable("FRONTIER_USERID") = RuleArguments("FRONTIER_USERID")
        SharedVariable("FRONTIER_PASSWORD") = RuleArguments("FRONTIER_PASSWORD")
        SharedVariable("FRONTIER_SYSTEM") = RuleArguments("FRONTIER_SYSTEM")
        SharedVariable("WORK_FOLDER") = RuleArguments("WORK_FOLDER")
        SharedVariable("HAVE_SUBMODELS") = RuleArguments("HAVE_SUBMODELS")
        SharedVariable("CAD_BOM_FILE") = RuleArguments("CAD_BOM_FILE")
        SharedVariable("A__SERVER_LOG") = RuleArguments("SERVER_LOG")
        SharedVariable("CADFLOW") = "True"

    Catch ex As Exception
        SharedVariable("SALES_ORDER") = "1"
        SharedVariable("SALES_ITEM") = "1"
        SharedVariable("WORK_ORDER") = "1"
        SharedVariable("CADFLOW_ACTION") = "JPG"
        SharedVariable("SERVER_INSTANCE") = "0"
        SharedVariable("ENVIRONMENT") = "FRNPCM031"
        SharedVariable("FRONTIER_USERID") = "PCM031"
        SharedVariable("FRONTIER_PASSWORD") = "PCM031"
        SharedVariable("FRONTIER_SYSTEM") = "S100EF60"
        SharedVariable("WORK_FOLDER") = "."
        SharedVariable("HAVE_SUBMODELS") = "N"
        SharedVariable("CAD_BOM_FILE") = ""
        SharedVariable("A__SERVER_LOG") = "C:/temp/ilogic.log"
        SharedVariable("CADFLOW") = "False"

    End Try
End Sub
