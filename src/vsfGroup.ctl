VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.UserControl vsfGroup 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000011&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   7815
   Begin VB.PictureBox picGroup 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1260
      TabIndex        =   0
      Tag             =   "Hello"
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VSFlex8UCtl.VSFlexGrid fg 
      Align           =   2  'Align Bottom
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   270
      Width           =   7815
      _cx             =   13785
      _cy             =   7435
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   12
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "vsfGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------
' API declarations

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long


'--------------------------------------------------------
' private declarations

Private Type POINTSGL
    X As Single
    Y As Single
End Type

Private Type GROUPINFO
    ctl As PictureBox
    text As String
End Type

Private Const CLR_BTNFACE = &H8000000F
Private Const CLR_BTNSHADOW = &H80000010
Private Const CLR_BTNHILITE = &H80000014

Private Const HELPMSG = "��ѡ������˳���ֶ�"
Private Const DRAG_TOLERANCE = 100 ' Twips

'--------------------------------------------------------
' variables

' mouse control
Private m_bCapture As Boolean   ' mouse captured?
Private m_bDragging As Boolean  ' dragging control?
Private m_ptDown As POINTSGL    ' where was the click
Private m_ptControl As POINTSGL ' original coordinates

Private m_iGroups As Integer    ' how many groups do we have
Private m_GroupInfo() As GROUPINFO ' group information vector

Private Function FindColumn(s$) As Integer
    
    ' locate column based on header text
    Dim i%
    For i = 0 To fg.Cols - 1
        If fg.Cell(flexcpTextDisplay, 0, i) = s Then
            FindColumn = i
            Exit Function
        End If
    Next
    
    ' this should never happen
    FindColumn = -1

End Function

Private Sub UpdateGrid()

    ' redraw is off to speed things up
    fg.Redraw = False
    
    ' move groups to left
    Dim i%, col%
    For i = 0 To m_iGroups - 1
        col = FindColumn(m_GroupInfo(i).text)
        fg.ColPosition(col) = i
    Next
    
    ' hide groups, make sure they're all sortable
    For i = 0 To m_iGroups - 1
        fg.ColHidden(i) = True
        If fg.ColSort(i) = 0 Then fg.ColSort(i) = flexSortGenericAscending
    Next
    
    ' show non-groups
    For i = m_iGroups To fg.Cols - 1
        fg.ColHidden(i) = False
    Next
    
    ' sort
    fg.Select fg.Row, 0, fg.Row, fg.Cols - 1
    fg.Sort = flexSortUseColSort

    ' create groups
    fg.Subtotal flexSTClear
    If m_iGroups > 0 Then
        
        For i = 0 To m_iGroups - 1
            fg.Subtotal flexSTNone, i, , , CLR_BTNFACE, , , , , True
        Next
        
        ' group them
       ' fg.Outline m_iGroups - 1
        fg.OutlineCol = m_iGroups
        fg.AutoSize m_iGroups
        
    End If
    
    ' move text to visible rows
    If m_iGroups > 0 Then
        For i = 1 To fg.Rows - 1
            If fg.IsSubtotal(i) Then
                Dim s$
                s = fg.Cell(flexcpTextDisplay, i, 0)
                fg.Cell(flexcpText, i, 0) = ""
                fg.Cell(flexcpText, i, m_iGroups) = s
            End If
        Next
    End If
    fg.MergeCells = flexMergeSpill

    ' redraw is back on
    fg.Redraw = True
    
End Sub


Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' if we clicked on a column, start dragging it
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then
        
        ' make sure we don't group on everything
        If m_iGroups >= fg.Cols - 1 Then
            Exit Sub
        End If
        
        ' which column are we grouping on?
        Dim col%
        col = fg.MouseCol
        
        ' confirm that this is a groupable column
        Dim i%
        For i = 0 To m_iGroups - 1
            If m_GroupInfo(i).text = fg.Cell(flexcpTextDisplay, 0, col) Then
                Cancel = True
                Beep
                Exit Sub
            End If
        Next
        ' UNDONE
        
        ' create entry in global array
        i = m_iGroups
        m_iGroups = m_iGroups + 1
        ReDim Preserve m_GroupInfo(i)
        
        ' create new group control
        Static newCtl%
        newCtl = newCtl + 1
        Load picGroup(newCtl)
        Set m_GroupInfo(i).ctl = picGroup(newCtl)
        m_GroupInfo(i).text = fg.Cell(flexcpTextDisplay, 0, col)
        
        ' init group control
        With picGroup(newCtl)
            .Tag = i
            .Width = .TextWidth(m_GroupInfo(i).text) + 2 * fg.RowHeight(0)
            .Height = fg.RowHeight(0) * 1.1
            .Move fg.ColPos(col), fg.top
            .Font = fg.Font
            .ZOrder
        End With
        
        ' save original position (none in this case)
        m_ptControl.X = -1
        m_ptControl.Y = -1
        
        ' start dragging
        m_bCapture = True
        m_bDragging = True
        m_ptDown.X = X - picGroup(newCtl).left
        m_ptDown.Y = fg.top + Y - picGroup(newCtl).top
        picGroup_Paint newCtl
        
        ' this is really cool:
        ' flex got the mouse down, but we want the group control to handle it
        ' so we set Cancel to true and transfer the mouse to the group control
        ' using the SetCapture API.
        Cancel = True
        With picGroup(newCtl)
            .Visible = True
            .SetFocus
            SetCapture .hwnd
        End With
    End If
End Sub

Private Sub picGroup_Click(Index As Integer)

    ' unless we were dragging, revert sort direction
    If (Not m_bDragging) And (m_ptControl.X > -1) Then
        
        ' revert sort direction
        Dim i%
        i = picGroup(Index).Tag
            fg.ColSort(i) = flexSortGenericAscending
        
        ' show the change
        UpdateLayout True
        
    End If
End Sub

Private Sub picGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    
    ' escape cancels dragging/clicking
    If (KeyAscii = 27) And (m_bCapture = True) Then
        
        ' move control back to its original position
        If m_bDragging Then
        
            ' if the group was still being created (not just dragged), delete it
            If m_ptControl.X < 0 And m_ptControl.Y < 0 Then
                DeleteGroup Index
            
            ' otherwise, move it back to where it was
            Else
                picGroup(Index).Move m_ptControl.X, m_ptControl.Y
            End If
        End If
        
        ' reset state variables
        m_bCapture = False
        m_bDragging = True
    
    End If
    
End Sub


Private Sub DeleteGroup(Index As Integer)
    
    ' remove control from the list
    Dim i%, J%
    i = picGroup(Index).Tag
    For J = i To m_iGroups - 2
        m_GroupInfo(J) = m_GroupInfo(J + 1)
    Next
    m_iGroups = m_iGroups - 1
    
    If m_iGroups = 0 Then fg.Outline 1

    ' hide/unload the control
    picGroup(Index).Visible = False
    If Index > 0 Then Unload picGroup(Index)
    
End Sub

Private Sub picGroup_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' left button starts dragging
    If Button = 1 Then
    
        ' save dragging information
        m_bCapture = True
        m_bDragging = False
        m_ptDown.X = X
        m_ptDown.Y = Y
        
        ' bring control to top, save its original position
        picGroup(Index).ZOrder
        m_ptControl.X = picGroup(Index).left
        m_ptControl.Y = picGroup(Index).top
    End If

End Sub

Private Sub picGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' drag control around
    If m_bCapture Then
        With picGroup(Index)
                        
            ' if we are not dragging yet, maybe it's time to start
            If Not m_bDragging Then
                If Abs(X - m_ptDown.X) > DRAG_TOLERANCE Then m_bDragging = True
                If Abs(Y - m_ptDown.Y) > DRAG_TOLERANCE Then m_bDragging = True
            End If
            
            ' if we're dragging, then do it
            If m_bDragging Then
            
                ' get new coordinates
                X = .left + (X - m_ptDown.X)
                Y = .top + (Y - m_ptDown.Y)
                
                ' restrict boundaries
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                If X > UserControl.ScaleWidth - .Width Then X = UserControl.ScaleWidth - .Width
                If Y > UserControl.ScaleHeight - .Height Then Y = UserControl.ScaleHeight - .Height
                If Y > fg.top Then Y = fg.top
            
                ' move the control
                .Move X, Y
                
                ' show where we'd go if we dropped now
                ' UNDONE
                
            End If
        End With
    End If
End Sub

Private Sub picGroup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' if we were dragging,
    ' we may have just moved the group to a new position, or
    ' we may have dropped it back into the grid
    If m_bDragging Then
        
        fg.Redraw = False
        
        ' back into grid, different position
        Y = picGroup(Index).top + Y
        If Y > fg.top Then
            
            ' see which column it was and where the mouse is
            Dim col%, i%
            col = FindColumn(m_GroupInfo(picGroup(Index).Tag).text)
            i = fg.MouseCol
            
            ' different? move column
            If i <> col Then
                fg.ColPosition(col) = i
            
            ' same? switch sort order
            Else
                If fg.ColSort(i) = flexSortGenericAscending Then
                    fg.ColSort(i) = flexSortGenericDescending
                Else
                    fg.ColSort(i) = flexSortGenericAscending
                End If
            End If
            
            ' remove our brand-new group
            DeleteGroup Index
        
        End If
        
        ' either way, show changes
        UpdateLayout True
        
        fg.Redraw = True
    End If

    ' cancel capture no matter what
    m_bCapture = False

End Sub


Private Sub picGroup_Paint(Index As Integer)
    
    Dim rc As RECT
    
    With picGroup(Index)
        
        ' draw frame
        rc.top = 0
        rc.left = 0
        rc.right = .Width / Screen.TwipsPerPixelX
        rc.bottom = .Height / Screen.TwipsPerPixelY
        DrawFrameControl .hDC, rc, DFC_BUTTON, DFCS_BUTTONPUSH
        
        ' draw text
        .CurrentX = .TextWidth(" ")
        .CurrentY = (.Height - .TextHeight(" ")) / 2.5
        picGroup(Index).Print m_GroupInfo(.Tag).text
        
        ' draw sort arrow if this is a group already
        If fg.ColWidth(.Tag) = 0 Then
            Dim X As Single, Y As Single, sz As Single
            sz = .Height * (1 / 3)
            X = .Width - sz
            
            ' pointing up
            If fg.ColSort(.Tag) = flexSortGenericDescending Then
                Y = (.Height - sz) / 2 + sz
                picGroup(Index).Line (X, Y)-(X - sz, Y), CLR_BTNHILITE
                picGroup(Index).Line -(X - sz / 2, Y - sz), CLR_BTNSHADOW
                picGroup(Index).Line -(X, Y), CLR_BTNHILITE
            
            ' pointing down
            Else
                Y = (.Height - sz) / 2
                picGroup(Index).Line (X, Y)-(X - sz, Y), CLR_BTNSHADOW
                picGroup(Index).Line -(X - sz / 2, Y + sz), CLR_BTNSHADOW
                picGroup(Index).Line -(X, Y), CLR_BTNHILITE
            End If
        End If
    End With

End Sub


Private Sub UserControl_Initialize()
    
    ' initialize embedded FlexGrid
    fg.SelectionMode = flexSelectionByRow
    fg.AllowUserResizing = flexResizeColumns
    fg.OutlineBar = flexOutlineBarComplete
    fg.ExplorerBar = flexExSortAndMove
    
    ' initialize group control based on grid data
    With picGroup(0)
        .Font = fg.Font
        .Height = fg.RowHeight(0)
        .Tag = 0
    End With

End Sub
Private Sub UpdateLayout(dogrid As Boolean)
    
    Dim swap As GROUPINFO
    Dim i%, cnt%, done%
    Dim X As Single, Y As Single, rh As Single
    Dim offsety As Single
    
    ' see how many groups are visible
    cnt = m_iGroups
    
    ' dimension and clear grouping area
    rh = fg.RowHeight(0)
    offsety = rh / 2
    Y = 2 * fg.RowHeight(0)
    If cnt > 1 Then Y = Y + (cnt - 1) * offsety
    Y = UserControl.ScaleHeight - Y
    If Y < 0 Then Y = 0
    fg.Height = Y
    UserControl.Cls
    
    ' if no groups, show helpful message
    If cnt = 0 Then
        UserControl.CurrentX = rh / 2
        UserControl.CurrentY = rh / 2
        UserControl.Print HELPMSG
        pxbl = HELPMSG
    Else
    pxbl = ""
    End If
    
    ' sort group vector by position (left-to-right)
    While Not done
        done = True
        For i = 0 To cnt - 2
            If m_GroupInfo(i).ctl.left > m_GroupInfo(i + 1).ctl.left Then
                done = False
                swap = m_GroupInfo(i)
                m_GroupInfo(i) = m_GroupInfo(i + 1)
                m_GroupInfo(i + 1) = swap
            End If
        Next
    Wend
    
    ' each control gets and index into the vector
    For i = 0 To cnt - 1
        m_GroupInfo(i).ctl.Tag = i
    Next
    
    ' position group controls
    Y = rh / 2
    X = Y
    For i = 0 To cnt - 1
        With m_GroupInfo(i).ctl
        
            ' move the control
            .Move X, Y
            Y = Y + offsety
            X = X + .Width + rh / 3
        
            ' draw connector
            If i < cnt - 1 Then
                UserControl.Line (X, Y + 2 / 3 * rh)-(X - rh * 2 / 3, Y + 2 / 3 * rh), 0
                UserControl.Line -(X - rh * 2 / 3, Y + rh / 2 - Screen.TwipsPerPixelY), 0
            End If
    
            ' draw placeholder
            UserControl.Line (.left, .top)-(.left + .Width - Screen.TwipsPerPixelX, .top + .Height - Screen.TwipsPerPixelY), 0, B
        
        End With
    Next
    
    ' redraw all controls at their new positions
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next
    UserControl.Refresh
    
    ' update the grid
    If dogrid Then UpdateGrid
    
    ' redraw all controls at their new positions (to show sort direction)
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next
    
End Sub

Private Sub UserControl_Resize()

    UpdateLayout False
    
End Sub



Public Property Get vsFlexGrid() As vsFlexGrid
    Set vsFlexGrid = fg
End Property


Public Sub Update()
    
    UpdateLayout True
    
End Sub

Public Sub px(Index As Integer)
    
    DeleteGroup (Index)
    picGroup_Click (Index)
End Sub

