'******************************************************************************
'* File:     comment2name.vbs
'* Title:    Name to Comment Conversion
'* Model:    Physical Data Model
'* Objects: Table, Column, View
'* Author:   luocheng
'* Created: 2016-5-10
'* Mod By:   
'* Modified: 
'* Version: 1.0
'* Memo:     Modify from name2comment.vbs
'******************************************************************************

Option   Explicit 
ValidationMode   =   True 
InteractiveMode   =   im_Batch

Dim   mdl   '   the   current   model

'   get   the   current   active   model 
Set   mdl   =   ActiveModel 
If   (mdl   Is   Nothing)   Then 
  MsgBox   "There   is   no   current   Model " 
ElseIf   Not   mdl.IsKindOf(PdPDM.cls_Model)   Then 
  MsgBox   "The   current   model   is   not   an   Physical   Data   model. " 
Else 
  ProcessFolder   mdl 
End   If

'   This   routine   copy   comment   into   name   for   each   table,   each   column   and   each   view 
'   of   the   current   folder 
Private   sub   ProcessFolder(folder) 
  Dim   Tab   'running     table 
  for   each   Tab   in   folder.tables 
    if   not   tab.isShortcut   then 
      tab.name   =   tab.name 
      Dim   col   '   running   column 
      for   each   col   in   tab.columns 
        if col.comment <> "" then
          col.name=   col.comment 
        elseif col.name <> "" then
          col.comment=   col.name
        
        end if
        
      next 
    end   if 
  next

  Dim   view   'running   view 
  for   each   view   in   folder.Views 
    if   not   view.isShortcut   then 
      view.name   =   view.comment 
    end   if 
  next

  '   go   into   the   sub-packages 
  Dim   f   '   running   folder 
  For   Each   f   In   folder.Packages 
    if   not   f.IsShortcut   then 
      ProcessFolder   f 
    end   if 
  Next 
end   sub
