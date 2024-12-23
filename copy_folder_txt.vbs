Option Explicit

Dim objFSO, objFolderA, objFolderB
Dim strFolderA, strFolderB

' フォルダAとフォルダBのパスを設定
strFolderA = "C:\\Users\\user\\Downloads\\sample_copy_folder_txt"  ' フォルダAのパスを指定
strFolderB = "C:\\Users\\user\\Downloads\\sample_copy_folder_txt_dst"  ' フォルダBのパスを指定

Set objFSO = CreateObject("Scripting.FileSystemObject")

' フォルダAが存在するか確認
If Not objFSO.FolderExists(strFolderA) Then
    WScript.Echo "フォルダAが見つかりません: " & strFolderA
    WScript.Quit
End If

' フォルダBが存在しない場合は作成
If Not objFSO.FolderExists(strFolderB) Then
    objFSO.CreateFolder strFolderB
End If

Set objFolderA = objFSO.GetFolder(strFolderA)
Set objFolderB = objFSO.GetFolder(strFolderB)

' フォルダA内のサブフォルダを処理
CopySubfoldersWithAEDT objFolderA, objFolderB

Sub CopySubfoldersWithAEDT(srcFolder, destFolder)
    Dim subFolder, destSubFolder

    ' サブフォルダをループ処理
    For Each subFolder In srcFolder.SubFolders
        ' コピー先のサブフォルダを作成
        If Not objFSO.FolderExists(destFolder.Path & "\\" & subFolder.Name) Then
            objFSO.CreateFolder destFolder.Path & "\\" & subFolder.Name
        End If
        Set destSubFolder = objFSO.GetFolder(destFolder.Path & "\\" & subFolder.Name)

        ' .txtファイルのみコピー
        Dim file
        For Each file In subFolder.Files
            If LCase(objFSO.GetExtensionName(file.Name)) = "txt" Then
                objFSO.CopyFile file.Path, destSubFolder.Path & "\\" & file.Name, True
            End If
        Next

        ' 再帰的にサブフォルダを処理
        CopySubfoldersWithAEDT subFolder, destSubFolder
    Next
End Sub