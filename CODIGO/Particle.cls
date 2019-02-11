VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Particle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Values Stored For Each Particle
Private mvarused          As Boolean
Private mvarsngX          As Single
Private mvarsngY          As Single
Private mvarsngXSpeed     As Single
Private mvarsngYSpeed     As Single
Private mvarsngXAccel     As Single
Private mvarsngYAccel     As Single
Private mvarsngR          As Single
Private mvarsngG          As Single
Private mvarsngB          As Single
Private mvarsngA          As Single
Private mvarsngAlphaDecay As Single

Public Sub ResetColor(sngRed As Single, _
                      sngGreen As Single, _
                      sngBlue As Single, _
                      sngAlpha As Single, _
                      sngDecay As Single)
    '************************************************************
    'Change the particle's color values
    'More info: http://www.vbgore.com/CommonCode.Particle.ResetColor
    '************************************************************
    
    On Error GoTo ResetColor_Err
    

    sngR = sngRed
    sngG = sngGreen
    sngB = sngBlue
    sngA = sngAlpha
    sngAlphaDecay = sngDecay

    
    Exit Sub

ResetColor_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "ResetColor"
    End If
Resume Next
    
End Sub

Public Sub ResetIt(X As Single, _
                   Y As Single, _
                   XSpeed As Single, _
                   YSpeed As Single, _
                   XAcc As Single, _
                   YAcc As Single)
    '************************************************************
    'Reset the particle's variables
    'More info: http://www.vbgore.com/CommonCode.Particle.ResetIt
    '************************************************************
    
    On Error GoTo ResetIt_Err
    

    sngX = X
    sngY = Y
    sngXSpeed = XSpeed
    sngYSpeed = YSpeed
    sngXAccel = XAcc
    sngYAccel = YAcc

    
    Exit Sub

ResetIt_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "ResetIt"
    End If
Resume Next
    
End Sub

Public Property Get sngA() As Single
    
    On Error GoTo sngA_Err
    

    sngA = mvarsngA

    
    Exit Property

sngA_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngA"
    End If
Resume Next
    
End Property

Public Property Let sngA(ByVal vData As Single)
    
    On Error GoTo sngA_Err
    

    mvarsngA = vData

    
    Exit Property

sngA_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngA"
    End If
Resume Next
    
End Property

Public Property Get sngAlphaDecay() As Single
    
    On Error GoTo sngAlphaDecay_Err
    

    sngAlphaDecay = mvarsngAlphaDecay

    
    Exit Property

sngAlphaDecay_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngAlphaDecay"
    End If
Resume Next
    
End Property

Public Property Let sngAlphaDecay(ByVal vData As Single)
    
    On Error GoTo sngAlphaDecay_Err
    

    mvarsngAlphaDecay = vData

    
    Exit Property

sngAlphaDecay_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngAlphaDecay"
    End If
Resume Next
    
End Property

Public Property Get sngB() As Single
    
    On Error GoTo sngB_Err
    

    sngB = mvarsngB

    
    Exit Property

sngB_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngB"
    End If
Resume Next
    
End Property

Public Property Let sngB(ByVal vData As Single)
    
    On Error GoTo sngB_Err
    

    mvarsngB = vData

    
    Exit Property

sngB_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngB"
    End If
Resume Next
    
End Property

Public Property Get sngG() As Single
    
    On Error GoTo sngG_Err
    

    sngG = mvarsngG

    
    Exit Property

sngG_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngG"
    End If
Resume Next
    
End Property

Public Property Let sngG(ByVal vData As Single)
    
    On Error GoTo sngG_Err
    

    mvarsngG = vData

    
    Exit Property

sngG_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngG"
    End If
Resume Next
    
End Property

Public Property Get sngR() As Single
    
    On Error GoTo sngR_Err
    

    sngR = mvarsngR

    
    Exit Property

sngR_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngR"
    End If
Resume Next
    
End Property

Public Property Let sngR(ByVal vData As Single)
    
    On Error GoTo sngR_Err
    

    mvarsngR = vData

    
    Exit Property

sngR_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngR"
    End If
Resume Next
    
End Property

Public Property Get sngX() As Single
    
    On Error GoTo sngX_Err
    

    sngX = mvarsngX

    
    Exit Property

sngX_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngX"
    End If
Resume Next
    
End Property

Public Property Let sngX(ByVal vData As Single)
    
    On Error GoTo sngX_Err
    

    mvarsngX = vData

    
    Exit Property

sngX_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngX"
    End If
Resume Next
    
End Property

Public Property Get sngXAccel() As Single
    
    On Error GoTo sngXAccel_Err
    

    sngXAccel = mvarsngXAccel

    
    Exit Property

sngXAccel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngXAccel"
    End If
Resume Next
    
End Property

Public Property Let sngXAccel(ByVal vData As Single)
    
    On Error GoTo sngXAccel_Err
    

    mvarsngXAccel = vData

    
    Exit Property

sngXAccel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngXAccel"
    End If
Resume Next
    
End Property

Public Property Get sngXSpeed() As Single
    
    On Error GoTo sngXSpeed_Err
    

    sngXSpeed = mvarsngXSpeed

    
    Exit Property

sngXSpeed_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngXSpeed"
    End If
Resume Next
    
End Property

Public Property Let sngXSpeed(ByVal vData As Single)
    
    On Error GoTo sngXSpeed_Err
    

    mvarsngXSpeed = vData

    
    Exit Property

sngXSpeed_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngXSpeed"
    End If
Resume Next
    
End Property

Public Property Get sngY() As Single
    
    On Error GoTo sngY_Err
    

    sngY = mvarsngY

    
    Exit Property

sngY_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngY"
    End If
Resume Next
    
End Property

Public Property Let sngY(ByVal vData As Single)
    
    On Error GoTo sngY_Err
    

    mvarsngY = vData

    
    Exit Property

sngY_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngY"
    End If
Resume Next
    
End Property

Public Property Get sngYAccel() As Single
    
    On Error GoTo sngYAccel_Err
    

    sngYAccel = mvarsngYAccel

    
    Exit Property

sngYAccel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngYAccel"
    End If
Resume Next
    
End Property

Public Property Let sngYAccel(ByVal vData As Single)
    
    On Error GoTo sngYAccel_Err
    

    mvarsngYAccel = vData

    
    Exit Property

sngYAccel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngYAccel"
    End If
Resume Next
    
End Property

Public Property Get sngYSpeed() As Single
    
    On Error GoTo sngYSpeed_Err
    

    sngYSpeed = mvarsngYSpeed

    
    Exit Property

sngYSpeed_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngYSpeed"
    End If
Resume Next
    
End Property

Public Property Let sngYSpeed(ByVal vData As Single)
    
    On Error GoTo sngYSpeed_Err
    

    mvarsngYSpeed = vData

    
    Exit Property

sngYSpeed_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "sngYSpeed"
    End If
Resume Next
    
End Property

Public Sub UpdateParticle(sngTime As Single)
    '************************************************************
    'Update the particle's variables
    'More info: http://www.vbgore.com/CommonCode.Particle.UpdateParticle
    '************************************************************
    
    On Error GoTo UpdateParticle_Err
    

    sngX = sngX + (LastOffsetX - ParticleOffsetX) + sngXSpeed * sngTime
    sngY = sngY + (LastOffsetY - ParticleOffsetY) + sngYSpeed * sngTime
    sngXSpeed = sngXSpeed + sngXAccel * sngTime
    sngYSpeed = sngYSpeed + sngYAccel * sngTime
    sngA = sngA - sngAlphaDecay * sngTime

    
    Exit Sub

UpdateParticle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "UpdateParticle"
    End If
Resume Next
    
End Sub

Public Property Get Used() As Boolean
    
    On Error GoTo Used_Err
    

    Used = mvarused

    
    Exit Property

Used_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "Used"
    End If
Resume Next
    
End Property

Public Property Let Used(ByVal vData As Boolean)
    
    On Error GoTo Used_Err
    

    mvarused = vData

    
    Exit Property

Used_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Particle" & "->" & "Used"
    End If
Resume Next
    
End Property

