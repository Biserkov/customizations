Private Sub Application_Startup()
    ' Adapted from https://www.slipstick.com/developer/code-samples/select-multiple-calendars-outlook/
    Dim navGroup As Outlook.NavigationGroup
    Dim NavFolder As Outlook.NavigationFolder
    Dim i As Integer
    
    Set Application.ActiveExplorer.CurrentFolder = Session.GetDefaultFolder(olFolderCalendar)
    DoEvents
    
    Set navGroup = Application.ActiveExplorer.NavigationPane.Modules.GetNavigationModule(olModuleCalendar).NavigationGroups.GetDefaultNavigationGroup(olMyFoldersGroup)

    For i = 1 To navGroup.NavigationFolders.Count
        Set NavFolder = navGroup.NavigationFolders.Item(i)
        NavFolder.IsSelected = True
        NavFolder.IsSideBySide = False
    Next

    Set navGroup = Nothing
    Set NavFolder = Nothing
End Sub
