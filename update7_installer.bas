Attribute VB_Name = "update7_installer"

Sub update7()
On Error GoTo handler

ActivePresentation.Slides(24).Shapes.AddOLEObject -20, 0, 10, 10, "Forms.Label.1"

Pausecode 1

Dim shp As Shape
For Each shp In ActivePresentation.Slides(24).Shapes.Range
    shp.Delete
Next

With ActivePresentation.Slides(24).Shapes.AddShape(msoShapeRectangle, 20, 50, 920, 50)
    .Line.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(50, 50, 50)
    With .TextFrame.TextRange
        .Font.Name = "Segoe UI"
        .Font.Bold = msoTrue
        .Font.Size = 16
        .Text = "Installing Maze Maker Update 1.4.0"
    End With
End With
With ActivePresentation.Slides(24).Shapes.AddShape(msoShapeRectangle, 20, 100, 920, 100)
    .Line.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(220, 220, 220)
    With .TextFrame.TextRange
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Color.RGB = RGB(20, 20, 20)
        .Text = "Please wait while we install the update. This action may just take a few seconds." & vbNewLine & "Please be patient and don't use your computer while the update is installing."
    End With
End With
With ActivePresentation.Slides(24).Shapes.AddShape(msoShapeRectangle, 20, 400, 920, 50)
    .Line.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(200, 200, 200)
End With
With ActivePresentation.Slides(24).Shapes.AddShape(msoShapeRectangle, 20, 400, 0.1 * 920, 50)
    .Name = "progress"
    .Line.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(50, 50, 50)
    With .TextFrame.TextRange
        .Font.Name = "Segoe UI Black"
        .Font.Size = 14
        .Font.Color.RGB = RGB(230, 230, 230)
        .Text = "10%"
    End With
End With
With ActivePresentation.Slides(24).Shapes.AddShape(msoShapeRectangle, 20, 350, 920, 20)
    .Name = "info"
    .Line.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(255, 255, 255)
    With .TextFrame.TextRange
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color.RGB = RGB(100, 100, 100)
        .Text = "Loading..."
    End With
End With

URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/loading.gif", ActivePresentation.Path & "\MazeMaker_Data\cache\update_loading.gif", 0, 0

With ActivePresentation.Slides(24).Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\cache\update_loading.gif", msoFalse, msoTrue, 380, 220)
    .Width = 200
End With

ActivePresentation.SlideShowWindow.View.GotoSlide 24
MsgBox "ola enzo"
updateprogress 0.05, "Loading..."


updateprogress 0.1, "Creating... '..\MazeMaker_Data\sprites'"
MkDir ActivePresentation.Path & "\MazeMaker_Data\sprites"

updateprogress 0.11, "Downloading... '..\MazeMaker_Data\sprites\colltime.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/colltime.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.png", 0, 0
updateprogress 0.12, "Downloading... '..\MazeMaker_Data\sprites\dottedblock.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/dottedblock.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\dottedblock.png", 0, 0
updateprogress 0.13, "Downloading... '..\MazeMaker_Data\sprites\finish_locked.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/finish_locked.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\finish_locked.png", 0, 0
updateprogress 0.14, "Downloading... '..\MazeMaker_Data\sprites\fullblock.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/fullblock.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.png", 0, 0
updateprogress 0.15, "Downloading... '..\MazeMaker_Data\sprites\goal.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/goal.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\goal.png", 0, 0
updateprogress 0.16, "Downloading... '..\MazeMaker_Data\sprites\off.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/off.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\off.png", 0, 0
updateprogress 0.17, "Downloading... '..\MazeMaker_Data\sprites\on.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/on.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\on.png", 0, 0
updateprogress 0.18, "Downloading... '..\MazeMaker_Data\sprites\onoff.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/onoff.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\onoff.png", 0, 0
updateprogress 0.19, "Downloading... '..\MazeMaker_Data\sprites\publish.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/publish.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\publish.png", 0, 0
updateprogress 0.2, "Downloading... '..\MazeMaker_Data\sprites\save.png'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/save.png", ActivePresentation.Path & "\MazeMaker_Data\sprites\save.png", 0, 0

updateprogress 0.25, "Downloading... '..\MazeMaker_Data\sprites\colltime.gif'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/colltime.gif", ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.gif", 0, 0
updateprogress 0.26, "Downloading... '..\MazeMaker_Data\sprites\fullblock.gif'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/fullblock.gif", ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.gif", 0, 0
updateprogress 0.27, "Downloading... '..\MazeMaker_Data\sprites\off.gif'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/off.gif", ActivePresentation.Path & "\MazeMaker_Data\sprites\off.gif", 0, 0
updateprogress 0.28, "Downloading... '..\MazeMaker_Data\sprites\on.gif'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/on.gif", ActivePresentation.Path & "\MazeMaker_Data\sprites\on.gif", 0, 0

updateprogress 0.33, "Downloading... '..\MazeMaker_Data\sfx\collect_time.wav'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/collect_time.wav", ActivePresentation.Path & "\MazeMaker_Data\sfx\collect_time.wav", 0, 0
updateprogress 0.36, "Downloading... '..\MazeMaker_Data\sfx\finish_unlock.wav'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/finish_unlock.wav", ActivePresentation.Path & "\MazeMaker_Data\sfx\finish_unlock.wav", 0, 0
updateprogress 0.37, "Downloading... '..\MazeMaker_Data\sfx\switch_blocks.wav'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/switch_blocks.wav", ActivePresentation.Path & "\MazeMaker_Data\sfx\switch_blocks.wav", 0, 0
updateprogress 0.39, "Downloading... '..\MazeMaker_Data\sfx\goal_change.wav'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/goal_change.wav", ActivePresentation.Path & "\MazeMaker_Data\sfx\goal_change.wav", 0, 0

updateprogress 0.41, "Downloading... '..\MazeMaker_Data\cache\update7\PPTGames_BetterArrays_v_1_20.cls'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/BetterArraysVBA/master/PPTGames_BetterArrays_v_1_20.cls", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\PPTGames_BetterArrays_v_1_20.cls", 0, 0

updateprogress 0.44, "Downloading... '..\MazeMaker_Data\cache\update7\PPTGames_VBAJSON_v_1_05.bas'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/PPTGames_VBAJSON_v_1_05.bas", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\PPTGames_VBAJSON_v_1_05.bas", 0, 0

updateprogress 0.48, "Importing... '" & Environ("systemroot") & "\system32\scrrun.dll" & "'"
ActivePresentation.VBProject.References.AddFromFile Environ("systemroot") & "\system32\scrrun.dll"

updateprogress 0.5, "Importing... '..\MazeMaker_Data\cache\update7\PPTGames_BetterArrays_v_1_20.cls'"
ActivePresentation.VBProject.VBComponents.Import ActivePresentation.Path & "\MazeMaker_Data\cache\update7\PPTGames_BetterArrays_v_1_20.cls"

updateprogress 0.55, "Importing... '..\MazeMaker_Data\cache\update7\PPTGames_VBAJSON_v_1_05.bas'"
ActivePresentation.VBProject.VBComponents.Import ActivePresentation.Path & "\MazeMaker_Data\cache\update7\PPTGames_VBAJSON_v_1_05.bas"

updateprogress 0.6, "Downloading... '..\MazeMaker_Data\cache\update7\editor'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/editor", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\editor", 0, 0

updateprogress 0.62, "Downloading... '..\MazeMaker_Data\cache\update7\player1'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/player1", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\player1", 0, 0

updateprogress 0.64, "Downloading... '..\MazeMaker_Data\cache\update7\player2'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/player2", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\player2", 0, 0

updateprogress 0.66, "Downloading... '..\MazeMaker_Data\cache\update7\online'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/online", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\online", 0, 0

updateprogress 0.68, "Downloading... '..\MazeMaker_Data\cache\update7\preview'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/preview", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\preview", 0, 0

updateprogress 0.7, "Downloading... '..\MazeMaker_Data\cache\update7\create'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/create", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\create", 0, 0

updateprogress 0.72, "Downloading... '..\MazeMaker_Data\cache\update7\saves'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/saves", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\saves", 0, 0

updateprogress 0.74, "Downloading... '..\MazeMaker_Data\cache\update7\publish'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/publish", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\publish", 0, 0

updateprogress 0.75, "Downloading... '..\MazeMaker_Data\cache\update7\changelog_en-gb.txt'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/changelog_en-gb", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\changelog_en-gb.txt", 0, 0

updateprogress 0.76, "Downloading... '..\MazeMaker_Data\cache\update7\changelog_pt-br.txt'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/changelog_pt-br", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\changelog_pt-br.txt", 0, 0

updateprogress 0.77, "Downloading... '..\MazeMaker_Data\cache\update7\changelog_pt-pt.txt'"
URLDownloadToFile 0, "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/changelog_pt-pt", ActivePresentation.Path & "\MazeMaker_Data\cache\update7\changelog_pt-pt.txt", 0, 0

updateprogress 0.8, "Importing... '..\MazeMaker_Data\cache\update7\editor'"
ActivePresentation.VBProject.VBComponents.item("Slide4").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide4").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide4").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\editor"

updateprogress 0.82, "Importing... '..\MazeMaker_Data\cache\update7\player1'"
ActivePresentation.VBProject.VBComponents.item("Slide5").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide5").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide5").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\player1"

updateprogress 0.84, "Importing... '..\MazeMaker_Data\cache\update7\player2'"
ActivePresentation.VBProject.VBComponents.item("Slide19").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide19").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide19").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\player2"

updateprogress 0.86, "Importing... '..\MazeMaker_Data\cache\update7\online'"
ActivePresentation.VBProject.VBComponents.item("Slide9").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide9").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide9").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\online"

updateprogress 0.88, "Importing... '..\MazeMaker_Data\cache\update7\preview'"
ActivePresentation.VBProject.VBComponents.item("Slide16").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide16").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide16").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\preview"

updateprogress 0.9, "Importing... '..\MazeMaker_Data\cache\update7\create'"
ActivePresentation.VBProject.VBComponents.item("Slide8").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide8").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide8").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\create"

updateprogress 0.91, "Importing... '..\MazeMaker_Data\cache\update7\saves'"
ActivePresentation.VBProject.VBComponents.item("Slide10").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide10").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide10").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\saves"

updateprogress 0.93, "Importing... '..\MazeMaker_Data\cache\update7\publish'"
ActivePresentation.VBProject.VBComponents.item("Slide18").CodeModule.DeleteLines 1, ActivePresentation.VBProject.VBComponents.item("Slide18").CodeModule.CountOfLines
ActivePresentation.VBProject.VBComponents.item("Slide18").CodeModule.AddFromFile ActivePresentation.Path & "\MazeMaker_Data\cache\update7\publish"

updateprogress 0.94, "Updating changelog files... '..\MazeMaker_Data\changelog\*'"
FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\changelog\en-gb.txt", Replace(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\cache\update7\changelog_en-gb.txt"), "\", "?")
FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\changelog\pt-br.txt", Replace(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\cache\update7\changelog_pt-br.txt"), "\", "?")
FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\changelog\pt-pt.txt", Replace(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\cache\update7\changelog_pt-pt.txt"), "\", "?")

updateprogress 0.96, "Updating language files... '..\MazeMaker_Data\langs\en-gb.mmlp'"
FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\langs\en-gb.mmlp", FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\langs\en-gb.mmlp") & vbNewLine & _
"""coins"":""coins"";" & vbNewLine & _
"""goal"":""Goal"";" & vbNewLine & _
"""goal_msg"":""Collect a least %coins% coins to unlock the finish of the maze."";" & vbNewLine & _
"""full_block"":""Full block"";" & vbNewLine & _
"""dotted_block"":""Dotted block"";" & vbNewLine & _
"""onoff_block"":""On/off block"";" & vbNewLine & _
"""collect_time"":""Collect time"";" & vbNewLine & _
"""online_warning_title"":""Important notice from PPTGames"";" & vbNewLine & _
"""online_warning"":""Thanks for installing the 1.4.0 update for Maze Maker!" & vbNewLine & vbNewLine & "Please, don't upload/publish mazes containing bugs or that are impossible to complete. We will remove mazes that don't follow these rules and may have to ban the author's account from upload/publishing mazes to the Maze Maker Online Community."";"

updateprogress 0.97, "Updating language files... '..\MazeMaker_Data\langs\pt-br.mmlp'"
FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\langs\pt-br.mmlp", FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\langs\pt-br.mmlp") & vbNewLine & _
"""coins"":""moedas"";" & vbNewLine & _
"""goal"":""Objetivo"";" & vbNewLine & _
"""goal_msg"":""Colete pelo menos %coins% moedas para desbloquear o final do labirinto."";" & vbNewLine & _
"""full_block"":""Bloco cheio"";" & vbNewLine & _
"""dotted_block"":""Bloco pontilhado"";" & vbNewLine & _
"""onoff_block"":""Bloco lig/des"";" & vbNewLine & _
"""collect_time"":""Coletar tempo"";" & vbNewLine & _
"""online_warning_title"":""Comunicado importante da PPTGames"";" & vbNewLine & _
"""online_warning"":""Obrigado por instalar a atualiza��o 1.4.0 do Maze Maker!" & vbNewLine & vbNewLine & "Por favor, n�o publique labirintos que contenham bugs ou que sejam imposs�veis de concluir. Removeremos labirintos que n�o seguem essas regras e talvez seja necess�rio proibir a conta do autor de publicar labirintos na comunidade online do Maze Maker."";"

updateprogress 0.98, "Updating language files... '..\MazeMaker_Data\langs\pt-pt.mmlp'"
FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\langs\pt-pt.mmlp", FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\langs\pt-pt.mmlp") & vbNewLine & _
"""coins"":""moedas"";" & vbNewLine & _
"""goal"":""Objetivo"";" & vbNewLine & _
"""goal_msg"":""Colete pelo menos %coins% moedas para desbloquear o final do labirinto."";" & vbNewLine & _
"""full_block"":""Bloco cheio"";" & vbNewLine & _
"""dotted_block"":""Bloco tracejado"";" & vbNewLine & _
"""onoff_block"":""Bloco lig/des"";" & vbNewLine & _
"""collect_time"":""Coletar tempo"";" & vbNewLine & _
"""online_warning_title"":""Comunicado importante da PPTGames"";" & vbNewLine & _
"""online_warning"":""Obrigado por instalar a atualiza��o 1.4.0 do Maze Maker!" & vbNewLine & vbNewLine & "Por favor, n�o publique labirintos que contenham bugs ou que sejam imposs�veis de concluir. Removeremos labirintos que n�o seguem estas regras e talvez seja necess�rio proibir a conta do autor de publicar labirintos na comunidade online do Maze Maker."";"

updateprogress 0.99, "Updating game content..."

With Slide5.Shapes.AddOLEObject(-100, 500, 70, 40, "WMPlayer.OCX.7")
    .Name = "wmp1"
    .OLEFormat.Object.settings.autoStart = False
End With

With Slide4.Shapes("game_txt_save")
    .TextFrame.TextRange.Text = ""
    .Width = .Height
    .Left = Slide4.Shapes("Ret�ngulo 88").Left - 2 * .Width - 26
End With

With Slide4.Shapes("game_txt_publish")
    .TextFrame.TextRange.Text = ""
    .Width = .Height
    .Left = Slide4.Shapes("Ret�ngulo 88").Left - .Width - 13
End With

With Slide4.Shapes.AddShape(msoShapeRectangle, Slide4.Shapes("game_txt_save").Left + Slide4.Shapes("game_txt_save").Width / 2 - 8, _
Slide4.Shapes("game_txt_save").Top + Slide4.Shapes("game_txt_save").Height / 2 - 8, 16, 16)
    .Line.Visible = msoFalse
    .Fill.UserPicture "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/save.png"
End With

With Slide4.Shapes.AddShape(msoShapeRectangle, Slide4.Shapes("game_txt_publish").Left + Slide4.Shapes("game_txt_publish").Width / 2 - 8, _
Slide4.Shapes("game_txt_publish").Top + Slide4.Shapes("game_txt_publish").Height / 2 - 8, 16, 16)
    .Line.Visible = msoFalse
    .Fill.UserPicture "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/publish.png"
End With

With Slide4.Shapes("mazecoloroption").Duplicate
    .Name = "savebtn"
    .Left = Slide4.Shapes("game_txt_save").Left
    .Top = Slide4.Shapes("game_txt_save").Top
    .Width = Slide4.Shapes("game_txt_save").Width
    .Height = Slide4.Shapes("game_txt_save").Height
    .ActionSettings.item(ppMouseClick).Run = "Slide4.save"
End With
Slide4.Shapes("game_txt_save").Name = "maze_save"

With Slide4.Shapes("mazecoloroption").Duplicate
    .Name = "publishbtn"
    .Left = Slide4.Shapes("game_txt_publish").Left
    .Top = Slide4.Shapes("game_txt_publish").Top
    .Width = Slide4.Shapes("game_txt_publish").Width
    .Height = Slide4.Shapes("game_txt_publish").Height
    .ActionSettings.item(ppMouseClick).Run = "Slide4.mazepublish"
End With
Slide4.Shapes("game_txt_publish").Name = "maze_publish"

With Slide4.Shapes.Range(Array("Conex�o reta 137", "mazetimeoption", "mazetime", "Gr�fico 16", "Ret�ngulo: Cantos Arredondados 73", _
"mazecoloroption", "Agrupar 78", "Ret�ngulo: Cantos Arredondados 76")).Group
    .Left = 600
    .Ungroup
End With

With Slide4.Shapes("Ret�ngulo: Cantos Arredondados 76").Duplicate
    .Name = "goalback"
    .Left = 512
    .Top = Slide4.Shapes("Ret�ngulo: Cantos Arredondados 76").Top
    .Width = 80
End With

With Slide4.Shapes("mazetime").Duplicate
    .Name = "mazegoal"
    .Left = 540
    .Top = Slide4.Shapes("mazetime").Top
    .Width = 50
    .TextFrame.TextRange.Text = "---"
End With

With Slide4.Shapes.AddShape(msoShapeRectangle, Slide4.Shapes("goalback").Left + 10, _
Slide4.Shapes("goalback").Top + Slide4.Shapes("goalback").Height / 2 - 8, 16, 16)
    .Line.Visible = msoFalse
    .Fill.UserPicture "https://raw.githubusercontent.com/PPTGames/MazeMakerUpdate7/master/goal.png"
End With

With Slide4.Shapes("mazecoloroption").Duplicate
    .Name = "goalbtn"
    .Left = Slide4.Shapes("goalback").Left
    .Top = Slide4.Shapes("goalback").Top
    .Width = Slide4.Shapes("goalback").Width
    .Height = Slide4.Shapes("goalback").Height
    .ActionSettings.item(ppMouseClick).Run = "Slide4.mazegoal"
End With

With Slide5.Shapes.AddShape(msoShapeRoundedRectangle, 355, 65, 250, 40)
    .Name = "goal_msg"
    .Fill.ForeColor.RGB = RGB(230, 192, 0)
    .Line.ForeColor.RGB = RGB(255, 255, 255)
    .Line.Weight = 2
    With .Shadow
        .Visible = msoTrue
        .Transparency = 0.8
        .Blur = 0
        .OffsetX = 0
        .OffsetY = -4
    End With
    With .TextFrame.TextRange
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 12
        .Font.Color.RGB = RGB(42, 35, 0)
        .Text = "Goal here"
    End With
End With

Slide4.Shapes("Ret�ngulo: Cantos Arredondados 90").Name = "elements_back"
Slide4.Shapes("elements_back").Width = 190
Slide4.Shapes.Range(Array("Agrupar 38", "item1option")).Group.Name = "temp_element_block_group"
With Slide4.Shapes("temp_element_block_group").Duplicate
    .Left = Slide4.Shapes("item2option").Left + Slide4.Shapes("item2option").Left - Slide4.Shapes("item1option").Left - 5
    .Top = Slide4.Shapes("item1option").Top
    .GroupItems(2).Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.png"
    .GroupItems(3).Name = "item3option"
    .Ungroup
End With
With Slide4.Shapes("temp_element_block_group").Duplicate
    .Left = Slide4.Shapes("item3option").Left + Slide4.Shapes("item2option").Left - Slide4.Shapes("item1option").Left - 5
    .Top = Slide4.Shapes("item1option").Top
    .GroupItems(2).Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\dottedblock.png"
    .GroupItems(3).Name = "item4option"
    .Ungroup
End With
With Slide4.Shapes("temp_element_block_group").Duplicate
    .Left = Slide4.Shapes("item4option").Left + Slide4.Shapes("item2option").Left - Slide4.Shapes("item1option").Left - 5
    .Top = Slide4.Shapes("item1option").Top
    .GroupItems(2).Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\off.png"
    .GroupItems(3).Name = "item5option"
    .Ungroup
End With
With Slide4.Shapes("temp_element_block_group").Duplicate
    .Left = Slide4.Shapes("item5option").Left + Slide4.Shapes("item2option").Left - Slide4.Shapes("item1option").Left - 5
    .Top = Slide4.Shapes("item1option").Top
    .GroupItems(2).Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.png"
    .GroupItems(3).Name = "item6option"
    .Ungroup
End With
Slide4.Shapes("temp_element_block_group").Ungroup
Slide4.Shapes("itemsel").ZOrder msoBringToFront
Slide4.Shapes.Range(Array("item1option", "item2option", "item3option", "item4option", "item5option", "item6option")).ZOrder msoBringToFront

Slide4.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.png", msoFalse, msoTrue, -100, 400, 40, 40).Name = "shp4"
Slide4.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\dottedblock.png", msoFalse, msoTrue, -100, 450, 40, 40).Name = "shp5"
Slide4.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\off.png", msoFalse, msoTrue, -100, 500, 40, 40).Name = "shp6"
Slide4.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.png", msoFalse, msoTrue, -100, 500, 40, 40).Name = "shp7"

Slide4.Shapes("shp4").ActionSettings.item(ppMouseClick).Action = ppActionRunMacro
Slide4.Shapes("shp4").ActionSettings.item(ppMouseClick).Run = "Slide4.shp_click"
Slide4.Shapes("shp5").ActionSettings.item(ppMouseClick).Action = ppActionRunMacro
Slide4.Shapes("shp5").ActionSettings.item(ppMouseClick).Run = "Slide4.shp_click"
Slide4.Shapes("shp6").ActionSettings.item(ppMouseClick).Action = ppActionRunMacro
Slide4.Shapes("shp6").ActionSettings.item(ppMouseClick).Run = "Slide4.shp_click"
Slide4.Shapes("shp7").ActionSettings.item(ppMouseClick).Action = ppActionRunMacro
Slide4.Shapes("shp7").ActionSettings.item(ppMouseClick).Run = "Slide4.shp_click"

With Slide5.Shapes.AddShape(msoShapeRectangle, -100, 500, 40, 40)
    .Name = "shp3"
    .Line.Visible = msoFalse
    .Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.gif"
End With
With Slide5.Shapes.AddShape(msoShapeRectangle, -100, 500, 40, 40)
    .Name = "shp5"
    .Line.Visible = msoFalse
    .Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\off.gif"
End With
Slide5.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.gif", msoFalse, msoTrue, -100, 500, 40, 40).Name = "shp6"

Slide5.Shapes("shp3").ActionSettings.item(ppMouseOver).Action = ppActionRunMacro
Slide5.Shapes("shp3").ActionSettings.item(ppMouseOver).Run = "Slide5.shp_hover"
Slide5.Shapes("shp5").ActionSettings.item(ppMouseClick).Action = ppActionRunMacro
Slide5.Shapes("shp5").ActionSettings.item(ppMouseClick).Run = "Slide5.onoff_click"
Slide5.Shapes("shp6").ActionSettings.item(ppMouseOver).Action = ppActionRunMacro
Slide5.Shapes("shp6").ActionSettings.item(ppMouseOver).Run = "Slide5.shp_hover"


With Slide19.Shapes.AddOLEObject(-100, 500, 70, 40, "WMPlayer.OCX.7")
    .Name = "wmp1"
    .OLEFormat.Object.settings.autoStart = False
End With


With Slide19.Shapes.AddShape(msoShapeRectangle, -100, 500, 40, 40)
    .Name = "shp3"
    .Line.Visible = msoFalse
    .Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.gif"
End With
With Slide19.Shapes.AddShape(msoShapeRectangle, -100, 500, 40, 40)
    .Name = "shp5"
    .Line.Visible = msoFalse
    .Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\off.gif"
End With

Slide19.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.gif", msoFalse, msoTrue, -100, 500, 40, 40).Name = "shp6"

Slide19.Shapes("shp3").ActionSettings.item(ppMouseOver).Action = ppActionRunMacro
Slide19.Shapes("shp3").ActionSettings.item(ppMouseOver).Run = "Slide19.shp_hover"
Slide19.Shapes("shp5").ActionSettings.item(ppMouseClick).Action = ppActionRunMacro
Slide19.Shapes("shp5").ActionSettings.item(ppMouseClick).Run = "Slide19.onoff_click"
Slide19.Shapes("shp6").ActionSettings.item(ppMouseOver).Action = ppActionRunMacro
Slide19.Shapes("shp6").ActionSettings.item(ppMouseOver).Run = "Slide19.shp_hover"

With Slide19.Shapes.AddShape(msoShapeRoundedRectangle, 355, 65, 250, 40)
    .Name = "goal_msg"
    .Fill.ForeColor.RGB = RGB(230, 192, 0)
    .Line.ForeColor.RGB = RGB(255, 255, 255)
    .Line.Weight = 2
    With .Shadow
        .Visible = msoTrue
        .Transparency = 0.8
        .Blur = 0
        .OffsetX = 0
        .OffsetY = -4
    End With
    With .TextFrame.TextRange
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 12
        .Font.Color.RGB = RGB(42, 35, 0)
        .Text = "Goal here"
    End With
End With

With Slide16.Shapes.AddShape(msoShapeRectangle, -100, 500, 40, 40)
    .Name = "shp3"
    .Line.Visible = msoFalse
    .Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\fullblock.gif"
End With
With Slide16.Shapes.AddShape(msoShapeRectangle, -100, 500, 40, 40)
    .Name = "shp5"
    .Line.Visible = msoFalse
    .Fill.UserPicture ActivePresentation.Path & "\MazeMaker_Data\sprites\off.gif"
End With
Slide16.Shapes.AddPicture(ActivePresentation.Path & "\MazeMaker_Data\sprites\colltime.gif", msoFalse, msoTrue, -100, 500, 40, 40).Name = "shp6"

With Slide16.Shapes.AddShape(msoShapeRoundedRectangle, 307.5, 145, 345, 70)
    .Name = "goal_msg"
    .Fill.ForeColor.RGB = RGB(230, 192, 0)
    .Line.ForeColor.RGB = RGB(255, 255, 255)
    .Line.Weight = 2
    With .Shadow
        .Visible = msoTrue
        .Transparency = 0.6
        .Size = 102
        .Blur = 10
        .OffsetX = 0
        .OffsetY = 4
    End With
    With .TextFrame.TextRange
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 30
        .Font.Color.RGB = RGB(42, 35, 0)
        .Text = "goal"
    End With
End With

Slide9.Shapes("refresh_btn").ActionSettings(ppMouseClick).Run = "refresh_btn"

Slide6.Shapes("collectedcoins").Visible = msoTrue
Slide7.Shapes("collectedcoins").Visible = msoTrue
Slide20.Shapes("collectedcoins").Visible = msoTrue
Slide21.Shapes("collectedcoins").Visible = msoTrue

With Slide9.Shapes.AddShape(msoShapeRoundedRectangle, -100, 500, 50, 40)
    .Name = "online_warning"
    .TextFrame.TextRange.Text = "0"
End With

ActivePresentation.Slides(4).SlideShowTransition.EntryEffect = ppEffectNone
ActivePresentation.Slides(5).SlideShowTransition.EntryEffect = ppEffectNone
ActivePresentation.Slides(19).SlideShowTransition.EntryEffect = ppEffectNone

updateprogress 0.99, "Saving game content..."
ActivePresentation.save

updateprogress 1, "Finishing update..."
Pausecode 2

With Slide17
    If .curupt.Caption = .lv.Caption Then
        FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\version.mmcf", Replace(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\version.mmcf"), _
        EQL5_GetElementValue(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\version.mmcf"), "code"), .lv.Caption)
        FS_ExportFile ActivePresentation.Path & "\MazeMaker_Data\version.mmcf", Replace(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\version.mmcf"), _
        EQL5_GetElementValue(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\version.mmcf"), "name"), "1.4.0")
        Slide2.Shapes("changelog").TextFrame.TextRange.Text = FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\changelog\" & game_getcursett("language") & ".txt")
        Slide2.Shapes("version").TextFrame.TextRange.Text = "ver " & EQL5_GetElementValue(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\version.mmcf"), "name")
        Slide17.Shapes("gamever").TextFrame.TextRange.Text = "ver " & EQL5_GetElementValue(FS_ReadFile(ActivePresentation.Path & "\MazeMaker_Data\version.mmcf"), "name")
        gamesettings_language_set game_getcursett("language")
        .Shapes("update_install").Visible = msoFalse
        .update_changelog.Visible = False
        .Shapes("update_finished").Visible = msoTrue
    End If
End With

ActivePresentation.SlideShowWindow.View.GotoSlide 17


Exit Sub

handler: MsgBox "UPDATE ERROR: " & error
game_createerrorlog "update7_installation" & vbNewLine & "Details:" & vbNewLine & Err & vbNewLine & error
End Sub

Sub updateprogress(value, Optional msg As String)
If msg <> "" Then ActivePresentation.Slides(24).Shapes("info").TextFrame.TextRange.Text = msg
ActivePresentation.Slides(24).Shapes("progress").TextFrame.TextRange.Text = (value * 100) & "%"
ActivePresentation.Slides(24).Shapes("progress").Width = value * 920
DoEvents
End Sub


