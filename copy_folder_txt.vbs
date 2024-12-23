Option Explicit

Dim objFSO, objFolderA, objFolderB
Dim strFolderA, strFolderB

' �t�H���_A�ƃt�H���_B�̃p�X��ݒ�
strFolderA = "C:\\Users\\user\\Downloads\\sample_copy_folder_txt"  ' �t�H���_A�̃p�X���w��
strFolderB = "C:\\Users\\user\\Downloads\\sample_copy_folder_txt_dst"  ' �t�H���_B�̃p�X���w��

Set objFSO = CreateObject("Scripting.FileSystemObject")

' �t�H���_A�����݂��邩�m�F
If Not objFSO.FolderExists(strFolderA) Then
    WScript.Echo "�t�H���_A��������܂���: " & strFolderA
    WScript.Quit
End If

' �t�H���_B�����݂��Ȃ��ꍇ�͍쐬
If Not objFSO.FolderExists(strFolderB) Then
    objFSO.CreateFolder strFolderB
End If

Set objFolderA = objFSO.GetFolder(strFolderA)
Set objFolderB = objFSO.GetFolder(strFolderB)

' �t�H���_A���̃T�u�t�H���_������
CopySubfoldersWithAEDT objFolderA, objFolderB

Sub CopySubfoldersWithAEDT(srcFolder, destFolder)
    Dim subFolder, destSubFolder

    ' �T�u�t�H���_�����[�v����
    For Each subFolder In srcFolder.SubFolders
        ' �R�s�[��̃T�u�t�H���_���쐬
        If Not objFSO.FolderExists(destFolder.Path & "\\" & subFolder.Name) Then
            objFSO.CreateFolder destFolder.Path & "\\" & subFolder.Name
        End If
        Set destSubFolder = objFSO.GetFolder(destFolder.Path & "\\" & subFolder.Name)

        ' .txt�t�@�C���̂݃R�s�[
        Dim file
        For Each file In subFolder.Files
            If LCase(objFSO.GetExtensionName(file.Name)) = "txt" Then
                objFSO.CopyFile file.Path, destSubFolder.Path & "\\" & file.Name, True
            End If
        Next

        ' �ċA�I�ɃT�u�t�H���_������
        CopySubfoldersWithAEDT subFolder, destSubFolder
    Next
End Sub