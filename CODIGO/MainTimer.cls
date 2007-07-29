VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

''
' The MainTimer of the game.
' Controls client´s intervals.
'
' @author Nacho (Integer) integer-x@hotmail.com
' @version 1.0
' @date 20060428

'*****************************************************
'Timer class for Argentum Online
'Designed and implemented by Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

Option Explicit

''
' How many timers we are going to use-
'
' @see See MainTimer.CreateAll
    Private Const CANTIDADTIMERS As Byte = 8

''
' A Timer data structure.
'
' @param Interval How long, in miliseconds, a cicle lasts.
' @param CurrentTick Current Tick in which the Timer is.
' @param StartTick Tick in which current cicle has started.
' @param Run True if the timer is active.

Private Type Timer
    Interval As Long
    CurrentTick As Long
    StartTick As Long
    Run As Boolean
End Type

'Timers
Dim Timer(1 To CANTIDADTIMERS) As Timer

''
' Timer´s Index.
'
' @param Attack                 Controls the Combat system.
' @param Work                   Controls the Work system.
' @param UseItemWithU           Controls the usage of items with the "U" key.
' @param UseItemWithDblClick    Controls the usage of items with double click.
' @param SendRPU                Controls the use of the "L" to request a pos update.
' @param CastSpell              Controls the casting of spells.
' @param Arrows                 Controls the shooting of arrows.
Public Enum TimersIndex
    Attack = 1
    Work = 2
    UseItemWithU = 3
    UseItemWithDblClick = 4
    SendRPU = 5
    CastSpell = 6
    Arrows = 7
    CastAttack = 8
End Enum


''
' Window´s API Function.
' A milisecond pricision counter.
'
' @return   Miliseconds since midnight.

Private Declare Function GetTickCount Lib "kernel32" () As Long

''
' Sets a new intrval for a timer.
'
' @param TimerIndex Timer´s Index
' @param Interval New lenght for the Timer´s cicle in miliseconds.
' @remarks  Must be donde after creating the timer and before using it, otherwise, Interval will be 0

Public Sub SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Long)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Sets a new interval for a timer.
'*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Interval = Interval
End Sub

''
' Starts the timer.
'
' @param TimerIndex Timer´s Index
' @remarks  Do this for starting the timer.

Public Sub Start(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Starts running the Timer.
'*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
End Sub

''
' Resume the timer form where we had stoped.
'
' @param TimerIndex Timer´s Index
' @remarks  Makes the Timer continue from where it had stoped.

Public Sub Continue(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Continue our timer from where we had stopped it.
'*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
    
    'We´ll set the StartTick again
    Timer(TimerIndex).StartTick = Timer(TimerIndex).StartTick + Timer(TimerIndex).CurrentTick
End Sub

''
' Pauses the Timer
'
' @param TimerIndex Timer´s Index
' @remarks  Can be resumed with Continue()

Public Sub Pause(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Pauses the timer.
'*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = False 'Don´t allow the timer to run.
    
    Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick 'Calculates acutual tick
End Sub

''
' Check if the timer has already completed it´s cicle.
'
' @param TimerIndex Timer´s Index
' @param Restart If true, restart if we timer has ticked
' @return True if the interval has already been elapsed
' @remarks  Can´t be done if the timer is stoped or if it had never been started.

Public Function Check(ByVal TimerIndex As TimersIndex, Optional Restart As Boolean = True) As Boolean
'*************************************************
'Author: Nacho Agustin (Integer)
'Last modified: 08/26/06
'Modification: NIGO: Added Restart as boolean
'Desc: Checks if the Timer has alredy "ticked"
'Returns: True if it has ticked, False if not.
'*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    If Not Timer(TimerIndex).Run Then Exit Function 'If we weren´t allowed to run, exit function
    
    Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick 'Calcutates CurrentTick
    
    If Timer(TimerIndex).CurrentTick >= Timer(TimerIndex).Interval Then
        Check = True 'We have Ticked!
        If Restart Then
            Timer(TimerIndex).StartTick = GetTickCount 'Restart Timer (Nicer than calling Restart() )
            If (TimerIndex = TimersIndex.Attack) Or (TimerIndex = TimersIndex.CastSpell) Then
                Timer(TimersIndex.CastAttack).StartTick = GetTickCount 'Set Cast-Attack interval
            ElseIf TimerIndex = TimersIndex.CastAttack Then
                Timer(TimersIndex.Attack).StartTick = GetTickCount 'Restart Attack interval
                Timer(TimersIndex.CastSpell).StartTick = GetTickCount 'Restart Magic interval
            End If
        End If
    End If
End Function

''
' Restarts timer.
'
' @param TimerIndex Timer´s Index

Public Sub Restart(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho Agustin (Integer)
'Last modified:
'Desc: Restarts timer
'*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).StartTick = GetTickCount
End Sub
