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
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

''
' The MainTimer of the game.
' Controls client´s intervals.
'
' @author Nacho (Integer) integer-x@hotmail.co,
' @version 1.0
' @date 20060428

'*****************************************************
'Timer class for Argentum Online
'Designed and implemented by Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

Option Explicit

'Timers
Dim Timer() As Timer

''
' Timer´s Index.
'
' @param Ataque Controls the Combat system.
Public Enum TimersIndex
    Ataque = 1
    Trabajo = 2
End Enum


''
' Window´s API Function.
' A milisecond pricision counter.
'
' @return   Miliseconds since midnight.
Private Declare Function GetTickCount Lib "kernel32" () As Long


''
' Creates all the Timers.
'
' @param Timers How many timers should be created.
' @remarks  Must be done before using Timers

Public Sub CreateAll(ByVal Timers As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Creates all the timers.
'*************************************************

'Re-Dimension them
ReDim Timer(1 To Timers)

'We set speed for all timers 1.
Dim Index As Integer
For Index = 1 To UBound(Timer)
    Timer(Index).Speed = 1
Next Index
End Sub

''
' Sets a new intrval for a timer.
'
' @param TimerIndex Timer´s Index
' @param Interval New lenght for the Timer´s cicle in miliseconds.
' @remarks  Must be donde after creating the timer and before using it, otherwise, Interval will be 0
Public Function SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Long)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Sets a new interval for a timer.
'*************************************************
Timer(TimerIndex).Interval = Interval
End Function

''
' Starts the timer.
'
' @param TimerIndex Timer´s Index
' @remarks  Do this for starting the timer.
Public Function Start(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Starts running the Timer.
'*************************************************
Timer(TimerIndex).Run = True 'We allow the Timer to run.

End Function

''
' Resume the timer form where we had stoped.
'
' @param TimerIndex Timer´s Index
' @remarks  Makes the Timer continue from where it had stoped.
Public Function Continue(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Continue our timer from where we had
'stopped it.
'*************************************************
Timer(TimerIndex).Run = True 'We allow the Timer to run.

'We´ll set the StartTick again
Timer(TimerIndex).StartTick = Timer(TimerIndex).StartTick + Timer(TimerIndex).CurrentTick
End Function

''
' Pauses the Timer
'
' @param TimerIndex Timer´s Index
' @remarks  Can be resumed with Continue()
Public Function Pause(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Pauses the timer.
'*************************************************
Timer(TimerIndex).Run = False 'Don´t allow the timer to run.

Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick 'Calculates acutual tick

End Function


''
' Check if the timer has already completed it´s cicle.
'
' @param TimerIndex Timer´s Index
' @remarks  Can´t be done if the timer is stoped or if it had never been started.
Public Function Check(ByVal TimerIndex As TimersIndex) As Boolean
'*************************************************
'Author: Nacho Agustin (Integer)
'Last modified:
'Desc: Checks if the Timer has alredy "ticked"
'Returns: True if it has ticked, False if not.
'*************************************************

If Not Timer(TimerIndex).Run Then Exit Function 'If we weren´t allowed to run, exit function

Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick 'Calcutates CurrentTick

'If CurrentTick (* Speed) is > than Interval  then,,,
If Timer(TimerIndex).CurrentTick * Timer(TimerIndex).Speed >= Timer(TimerIndex).Interval Then

    Check = True 'We have Ticked!
    Restart (TimerIndex) 'Restart Timer
End If

End Function

''
' Restarts timer.
'
' @param TimerIndex Timer´s Index
Public Function Restart(ByVal TimerIndex As TimersIndex)
'*************************************************
'Author: Nacho Agustin (Integer)
'Last modified:
'Desc: Restarts timer
'*************************************************
Timer(TimerIndex).StartTick = GetTickCount
End Function

''
' Sets speed for the timer.
'
' @param TimerIndex Timer´s Index
' @param Wished speed for the Timer.
' @remarks  Speed is, by default, 1
Public Function SetSpeed(ByVal TimerIndex As TimersIndex, ByVal Speed As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified:
'Desc: Set Timer's speed
'*************************************************
Timer(TimerIndex).Speed = Speed
End Function

