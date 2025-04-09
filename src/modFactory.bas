Attribute VB_Name = "modFactory"
Option Explicit

'/**
' * @description �g��������N���X�̃C���X�^���X���쐬����t�@�N�g���֐�
' * @param {String} pValue - ���������镶����l
' * @param {Boolean} [pMutable=False] - �~���[�^�u������������邩�ǂ���
' * @returns {clsEnhancedString} ���������ꂽ�g��������N���X�̃C���X�^���X
' * @example
' * Dim str As clsEnhancedString
' * Set str = CreateEnhancedString("Hello", True)
' * Debug.Print str.ToUpperCase().Concat(", World!").Value ' �o��: "HELLO, World!"
' */
Public Function CreateEnhancedString(ByVal pValue As String, Optional ByVal pMutable As Boolean = False) As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = New clsEnhancedString
    
    lvResult.Initialize pValue, pMutable
    
    Set CreateEnhancedString = lvResult
End Function
