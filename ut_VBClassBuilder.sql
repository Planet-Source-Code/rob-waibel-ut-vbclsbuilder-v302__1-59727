Drop Proc ut_VBClassBuilder
GO
Create Proc ut_VBClassBuilder
  @TableName VarChar(50)
AS

/************************************************************************
*****  Created By Robert Waibel, Phoenix, AZ 03/28/2005             
*****  ut_VBClassBuilder Version 3.0.2
*****
*****       This Utility Procedure is designed to assist a VB or VBA
*****  developer in creating classes based on SQL Server Databases. 
*****  There are two Functions that are called in this created code,
*****  they are SQLData() and InfoChanged.  These have also been    
*****  submitted by me.                                             
*****                                                               
*****  Revisions                                                    
*****  	Found an issue with the Criteria Properties 3/28/2005
*****	Added ValidateClass to the save subroutine  3/29/2005
*****  
*****  USE:
*****  	ut_VBClassBuilder 'dtProperties'
*****  
*****  
*****  
*****  
************************************************************************/
  Declare @ID INT 
  Declare @@ColName VarChar(24), @@Type INT, @@FieldType VarChar(128), @@FLength INT, @@Nullable TinyInt, @@Default TinyInt
  Declare @SQLSelect VarChar(4000), @LoadRec1 VarChar(2000), @SQL5 VarChar(2000)
  Declare @ClassType Varchar(2000), @ClassStart VarChar(2000)
  Declare @LoadRec VarChar(2000), @SQL6 VarChar(2000)
  Declare @ScriptMove as VarChar(4000), @SQLArray Varchar(4000), @SQLSave1 VarChar(8000)
  DECLARE @SQLSave2 VarChar(8000), @SQLSave3 Varchar(8000), @LoadClass3 Varchar(2000)
  DECLARE @PropertyALL VarChar(8000), @SQLLen INT, @LenMax INT, @VBType VarChar(16)
	DECLARE @CriteraSQL Varchar(4000)
	Declare @@PKField sysname
	Declare @LoadClass VarChar(128)
	DECLARE @full_table_name	nvarchar(255)
	Declare @ClassValidation as varchar(4000)

  Select @LenMax = 100
  Select @ID = object_id(quotename(@TableName)) 
  if @ID is null
    Return -1
  else Begin
	EXEC sp_fkeysRDW @TableName, NULL
	EXEC sp_fkeysRDW NULL, @TableName



	Select @full_table_name = quotename(@TableName)
    	Declare Z Cursor For 
     	Select SC.Xtype ,SC.name 'FieldName', st.name 'FieldType', SC.Length, isnullable, 
			case when cDefault > 0 then 1 else 0 end 'ServerDefault'
        		From SysCOlumns SC  INNER JOIN SysTypes ST on SC.xtype = St.xtype
      	where ID = @ID
			Order By ColOrder
	For Read Only

    Open Z

	Print 'Option Explicit'
	Print ''
	Print '''NOTE: errMissingRequiredInfo is a Custom Error Number'
	Print '''const errMissingRequiredInfo = bnObjectError + 513'

    Select @SQLSelect = '', @ClassStart = ''
	SELECT @ClassValidation = 'Private Function ValidateClass() as boolean' + char(13) + char(10) +
		space(5) + 'Dim X as integer' + char(13) + char(10) + 
		space(5) + 'ValidateClass = true' + char(13) + char(10) +
		space(5) + 'X = 0' + char(13) + char(10) +
		space(5) + 'Redim InValid(X)' + char(13) + char(10)

    SELECT @SQLArray = char(13) + char(10)
    SELECT @SQLArray = @SQLArray + 'Public Sub FillArray(inARRAY() as Variant)' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(5) + 'ForArray = true' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(5) + 'LoadClass' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(5) + 'if not RS.EOF Then' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(10) + 'inarray = rs.getrows' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(5) + 'else' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(10) + 'redim inArray(rs.fields.count-1,0)' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(10) + 'inArray(0,0) = "NONE"' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + space(5) + 'End If' + char(13) + char(10)
    SELECT @SQLArray = @SQLArray + 'End Sub' + char(13) + char(10)

    SELECT @SQLSave1 = char(13) + char(10)
    	SELECT @SQLSave1 = @SQLSave1 + 'Public sub Save()' + char(13) + char(10)
    		+ space(5) + 'dim iSQL as string, vSQL as string' + char(13) + char(10)
    		+ space(5) + 'if IsNewRec then' + char(13) + char(10)
		+ space(10) + 'if ValidateClass then' + char(13) + char(10)
		+ space(15) + 'iSQL = "INSERT INTO ' + @TAbleName + '(' + char(13) + char(10)
    		+ space(15) + 'vsql = " Values ("' + char(13) + char(10)

    	Select @SQLSave2 = space(10) + 'Else' + char(13) + char(10) 
		+ space(15) + 'Dim i As Integer, msg As String' + char(13) + char(10) 
          + space(15) + 'msg = "You are missing information for the following:"' + char(13) + char(10)
          + space(15) + 'For i = 0 To UBound(InValid)' + char(13) + char(10)
          + space(20) + 'msg = msg & vbCrLf & InValid(i)' + char(13) + char(10)
          + space(15) + 'Next i' + char(13) + char(10)
          + space(15) + 'msg = msg & vbCrLf & "Please enter the information before continuing..."' + char(13) + char(10)
          + space(15) + 'Err.Raise errMissingRequiredInfo, "cls' + @TableName + '", msg' + char(13) + char(10)
          + space(15) + 'Exit Sub' + char(13) + char(10)
        	+ space(10) + 'End If' + char(13) + char(10)
		+ space(05) + 'ELSE' + char(13) + char(10) + space(08) + 'isql = "": vsql = ""' + char(13) + char(10)

	Select @ClassStart = @ClassStart + char(13) + char(10)
    Select @ClassStart = @ClassStart + 'Dim RS As New ADODB.Recordset' + char(13) + char(10)
    Select @ClassStart = @ClassStart + 'Dim sSQL As String' + char(13) + char(10)
    Select @ClassStart = @ClassStart + '' + char(13) + char(10)
    Select @ClassStart = @ClassStart + 'Dim ForArray As Boolean' + char(13) + char(10)
    Select @ClassStart = @ClassStart + 'Dim IsNewRec As Boolean' + char(13) + char(10)
    Select @ClassStart = @ClassStart + 'Dim InValid() as Variant ''used to trap invalid field values' + char(13) + char(10)
    Select @ClassStart = @ClassStart + '' + char(13) + char(10)
	Select @ClassStart = @ClassStart + 'dim sCriteria as string' + char(13) + char(10)
    Select @ClassStart = @ClassStart + '' + char(13) + char(10)

    select @ClassType = 'Private Type typ' + @TableName + char(13) + char(10)
  
    Select @LoadRec = 'Private SUB LoadRec' + char(13) + char(10)
    Select @LoadRec = @LoadRec + space(5) + 'If not rs.eof then' + char(13) + char(10)
    Select @LoadRec1 = space(5) + 'Else' + char(13) + char(10)

    Select @Sql6 = 'Private Sub Class_Initialize()' + char(13) + char(10)
    Select @SQL6 = @SQL6 + space(05) + 'If cnCLS is Nothing Then Set cnCLS = ADOConnect' + char(13) + char(10)
    Select @SQL6 = @SQL6 + 'End Sub' + char(13) + char(10)
    Select @SQL6 = @SQL6 + '' + char(13) + char(10)
    Select @SQL6 = @SQL6 + 'Private Sub Class_Terminate()' + char(13) + char(10)
    Select @SQL6 = @SQL6 + space(05) + 'Set RS = Nothing' + char(13) + char(10)
    Select @SQL6 = @SQL6 + 'End Sub' + char(13) + char(10)

--*****  Get The Primary Key
	Declare tmp CURSOR For 
		select convert(sysname,c.name)
			from sysindexes i, syscolumns c, sysobjects o 
		where	o.id = @ID
				and o.id = c.id
				and o.id = i.id
				and (i.status & 0x800) = 0x800
				and (c.name = index_col (@full_table_name, i.indid,  1) or
		    		c.name = index_col (@full_table_name, i.indid,  2) or
		    		c.name = index_col (@full_table_name, i.indid,  3) or
		    		c.name = index_col (@full_table_name, i.indid,  4) or
		    		c.name = index_col (@full_table_name, i.indid,  5) or
		    		c.name = index_col (@full_table_name, i.indid,  6) or
		     	c.name = index_col (@full_table_name, i.indid,  7) or
		     	c.name = index_col (@full_table_name, i.indid,  8) or
		     	c.name = index_col (@full_table_name, i.indid,  9) or
		    		c.name = index_col (@full_table_name, i.indid, 10) or
		    		c.name = index_col (@full_table_name, i.indid, 11) or
		     	c.name = index_col (@full_table_name, i.indid, 12) or
		     	c.name = index_col (@full_table_name, i.indid, 13) or
		     	c.name = index_col (@full_table_name, i.indid, 14) or
		    		c.name = index_col (@full_table_name, i.indid, 15) or
		    		c.name = index_col (@full_table_name, i.indid, 16)
		   		)
			order by 1
		Open tmp
		Fetch Next From tmp into @@PKField
		Deallocate tmp

    select @ID = 0, @PropertyALL = ''
    Select @SQLLen = Len(@SQLSelect)
/*****  Cursor Run starts Here *****/
    Fetch Next From Z Into @@Type, @@ColName, @@FieldType , @@FLength , @@Nullable, @@Default

    While @@Fetch_Status = 0 Begin
--*****  Group the SQL Xtypes into VB Types
		Select @VBType = 
			case when @@type = 34 then ' Image'
				when @@type = 35 then ' String'
				when @@type = 36 then ' Variant'
				when @@type = 48 then ' Integer'
				when @@type = 52 then ' Small'
				when @@type = 56 then ' Long'
				when @@type = 58 and @@Nullable = 0 then ' Date'
				when @@type = 58 and @@Nullable = 1 then ' Variant'
				when @@type = 59 then ' Double'
				when @@type = 60 then ' Currency'
				when @@type = 61 and @@Nullable = 0 then ' Date'
				when @@type = 61 and @@Nullable = 1 then ' Variant'
				when @@type = 62 then ' Double'
				when @@type = 98 then ' Variant'
				when @@type = 99 then ' String'
				when @@type = 104 then ' Boolean'
				when @@type = 106 then ' Double'	
				when @@type = 108 then ' Double'
				when @@type = 122 then ' Currency'
				when @@type = 127 then ' Long'
				when @@type = 165 then ' Image'
				when @@type = 167 then ' String'
				when @@type = 173 then ' Image'	
				when @@type = 175 then ' String'
				when @@type = 189 then ' Variant'
				when @@type = 231 then ' String'	
				when @@type = 239 then ' String'
			END
				
/*****  Define the Property Calls Here *****/
--*****  Property Let
     Select @PropertyALL = @PropertyALL + 'Public Property Let'
     Select @PropertyALL = @PropertyALL + ' ' + rtrim(@@ColName) + '(' + lower(substring(@VBType, 2,1)) + 'Data as' + @VBType
		+ ')' + char(13) + char(10)
	if @@PKField = @@ColName 
		Select @PropertyALL = @PropertyALL + Space(5) + '''Primary KEY' + char(13) + char(10)
     Select @PropertyALL = @PropertyALL + space(5) + 'Rec.' + rtrim(@@ColName) + ' = '+ lower(substring(@VBType, 2,1)) + 'Data' + char(13) + char(10)
     Select @PropertyALL = @PropertyALL + 'End Property' + char(13) + char(10) + char(13) + char(10)

--*****  Property Get
      Select @PropertyALL = @PropertyALL + 'Public Property Get' + ' ' + rtrim(@@ColName) + '() as' + @VBType         
		+ char(13) + char(10)
	if @@PKField = @@ColName
     	Select @PropertyALL = @PropertyALL + Space(5) + '''Primary KEY' + char(13) + char(10)
      Select @PropertyALL = @PropertyALL + space(5) + rtrim(@@ColName) + ' = ' + 'Rec.' + rtrim(@@ColName) + char(13) + char(10)
      Select @PropertyALL = @PropertyALL + 'End Property' + char(13) + char(10) + char(13) + char(10)

/*****  Define SQL Statement and function *****/      
      if @SqlLen + Len(rtrim(@@ColName)) >= @LenMax
        Begin
          select @SQLSelect = @SQLSelect + '" _' + char(13) + char(10) + space(10) + '& "'
          select @LenMax = @LenMax + 100
        END
      Select @SQLSelect = @SQLSelect + rtrim(@@ColName) + ', '
      select @sqlLen = @SqlLen + Len(rtrim(@@ColName))

/*****  Define the Class Type *****/
      Select @ClassType = @ClassType + space(05) + rtrim(@@ColName) + ' as' + @VBType
	if @@PKField = @@ColName 
		Select @ClassType = @ClassType + Space(5) + '''<-- Primary KEY'

	Select @ClassType = @ClassType + char(13) + char(10)

/*****  Load Routine *****/
    Select @LoadRec = @LoadRec + space(10) + 'Rec.'+ rtrim(@@ColName) + 
		' = RS("'+ rtrim(@@ColName) +'")' + 
	Case when @@Nullable = 1 and (@@Type not in (58, 61)) then ' & ""' Else '' END + char(13) + char(10)

    Select @LoadRec1 = @LoadRec1 + space(10) + 'Rec.' + rtrim(@@ColName) + ' = ' +
      (Case When (@@Type in (48, 52, 56, 59, 60, 62, 106, 108, 122))  THEN ' 0'
            When (@@Type in (36, 58, 61, 98, 165, 173, 189)) then ' NULL'
            When @@Type = 104 then ' False'            
            Else ' ""' 
        END) + char(13) + char(10)

/*****  The Save Routine Start and cycle *****/
    Select @SQLSave1 = @SQLSave1 + space(15) + 'if' +
	      (Case When (@@Type in (48, 52, 56, 59, 60, 62, 106, 108, 122)) then ' Rec.' + rtrim(@@ColName) + ' > 0'
            When (@@Type in (35, 36, 58, 61, 98, 99, 165, 167, 173, 189, 231, 239)) then ' Len(Rec.' + rtrim(@@ColName) + ') > 0'
            When @@Type = 104 then ' False'            
            Else ' len(Rec.' + rtrim(@@ColName) + ') > 0'
        END) + ' then' + char(13) + char(10)
    Select @SQLSave1 = @SQLSave1 + space(20) + 'isql = isql & "' + rtrim(@@ColName)+ ', " ' + char(13) + char(10)
    Select @SQLSave1 = @SQLSave1 + space(20) + 'vsql = vsql & ' + 
	(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then 'sqldata(rec.'
            ELSE 'rec.'            
            END) + rtrim(@@ColName) + 
	(Case  When (@@Type in (35, 99, 167, 175, 231, 239)) then ', vbstring)'
            When (@@Type in (58, 61)) then ', vbdate)'
            When @@Type = 104 then ', vbBoolean)'            
            ELSE ''
            END) + ' & ", "' + char(13) + char(10)
    Select @SQLSave1 = @SQLSave1 + space(15) + 'End if' + char(13) + char(10)
	if @@PKField <> @@ColName 
      Begin
        Select @SQLSAVE2 = @SQLSave2 + space(08) + 'if InfoChanged(Rec.' + rtrim(@@ColName)+ ', Orig.'+ rtrim(@@ColName) +') THEN isql = isql & "'
        Select @SQLSAVE2 = @SQLSave2 + rtrim(@@ColName) + ' = " & ' + 
			(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then 'sqldata(rec.'
            ELSE 'rec.'            
            END) + rtrim(@@ColName) + 
	(Case  When (@@Type in (35, 99, 167, 175, 231, 239)) then ', vbstring)'
            When (@@Type in (58, 61)) then ', vbdate)'
            When @@Type = 104 then ', vbBoolean)'            
            ELSE ''
            END) + ' & ", "' + char(13) + char(10)
      END -- if @@PKField <> @@ColName 
    else
      Begin
        Select @SQLSave3 = + space(15) + 'isql = isql & " WHERE ' + rtrim(@@ColName) + ' = " & '+ 
		(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then 'sqldata(rec.'
          	ELSE 'rec.'            
            	END) + rtrim(@@ColName) + 
		(Case  	When (@@Type in (35, 99, 167, 175, 231, 239)) then ', vbstring)'
            		When (@@Type in (58, 61)) then ', vbdate)'
            		When @@Type = 104 then ', vbBoolean)'            
           	ELSE ''
            	END) + ' & ", "' + char(13) + char(10)

      END -- Else if @@PKField = @@ColName 

/*****  The Sub LoadClass*****/
    Select @ID = @ID + 1
	if @@PKField = @@ColName BEGIN
		select @LoadClass = 'Public Sub LoadClass (Optional pk' + rtrim(@@ColName) + ' as' + @VBType + ' = ' +
			(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then '""' 
				ELSE '0'
				END) + ')' + char(13) + char(10) 
	    	Select @LoadClass3 = Space(5) + 'If ' +
			(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then 'Len(' 
				ELSE ''
				END) + 'pk'+ rtrim(@@ColName) + 
			(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then ') > ' 
				ELSE ' >'
				END) + ' 0 then' + char(13) + char(10) +
			space(10) + 'ssql = ssql & " WHERE ' + rtrim(@@ColName) + ' = " & ' + 
			(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then 'sqldata(pk'
     	     	ELSE 'pk'            
          	  	END) + rtrim(@@ColName) + 
			(Case  	When (@@Type in (35, 99, 167, 175, 231, 239)) then ', vbstring)'
     	       		When (@@Type in (58, 61)) then ', vbdate)'
          	  		When @@Type = 104 then ', vbBoolean)'            
           		ELSE ''
	            	END) + char(13) + char(10) + Space(5) +
			'Else' + char(13) + char(10) +
			Space(10) + 'If len(sCriteria) > 0 then ssql = ssql & " WHERE " & sCriteria' + char(13) + char(10) +
			space(5) + 'End If' + char(13) + char(10)


	END --if @@PK
--****  Populate Class Validation
	if @@Nullable = 0 BEGIN
		if @@Default = 0 BEGIN
			SELECT @ClassValidation = @ClassValidation +
				space(05) + 'if ' +
				(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then 'isNULL(' 
					ELSE ''
					END) + 'Rec.'+ rtrim(@@ColName) + 
				(Case  When (@@Type not in (48, 52, 56, 59, 60, 62, 106, 108, 122, 127, 165, 173)) then ') > ' 
					ELSE ' = '
					END) + ' 0 then' + char(13) + char(10) +
				space(10) + 'X = X + 1' + char(13) + char(10) +
				space(10) + 'Redim Preserve inValid(X)' + char(13) + char(10) +
				space(10) + 'Invalid(x) = "' + rtrim(@@ColName) + '"' + char(13) + char(10) +
				space(05) + 'End if' + char(13) + char(10)
		END
	END


    Fetch Next From Z Into @@Type, @@ColName, @@FieldType , @@FLength , @@Nullable, @@Default
  End --  While 
  Close Z
  Deallocate Z
/*****  the SAVE routine *****/
    Select @SQLSAVE1 = @SQLSave1 + space(15) + 'if len(isql) > 0 then' + char(13) + char(10)
    Select @SQLSAVE1 = @SQLSave1 + space(20) + 'isql = left(isql, len(isql)-2) & ")"' + char(13) + char(10)
    Select @SQLSAVE1 = @SQLSave1 + space(20) + 'vsql = left(vsql, len(vsql)-2) & ")"' + char(13) + char(10)
    Select @SQLSAVE1 = @SQLSave1 + space(20) + 'cnCLS.execute isql & vsql' + char(13) + char(10)
    Select @SQLSAVE1 = @SQLSave1 + space(15) + 'end if' 


    Select @SQLSAVE2 = @SQLSave2 + space(10) + 'If Len(isql) > 0 then'  + char(13) + char(10)
    Select @SQLSAVE2 = @SQLSave2 + space(15) + 'isql = "Update ' + @TableName + ' SET " & left(isql, len(isql)-2)' + char(13) + char(10)
    Select @SQLSAVE2 = @SQLSave2 + @SQLSave3
    Select @SQLSAVE2 = @SQLSave2 + space(15) + 'cnCLS.execute isql' + char(13) + char(10)
    Select @SQLSAVE2 = @SQLSave2 + space(10) + 'end if' + char(13) + char(10)
    Select @SQLSAVE2 = @SQLSave2 + space(05) + 'end if' + char(13) + char(10)
    Select @SQLSAVE2 = @SQLSave2 + space(05) + 'Orig = Rec' + char(13) + char(10)
    Select @SQLSAVE2 = @SQLSave2 + 'End Sub' + char(13) + char(10)

/*****  Finish the Class Type Definitions *****/
    select @ClassType = @ClassType + 'End Type'+ char(13) + char(10)
    select @ClassType = @ClassType + 'Dim Rec as typ'+ @TableName + char(13) + char(10)
    select @ClassType = @ClassType + 'Dim Orig as typ'+ @TableName + char(13) + char(10)

/*****  Allow the Recordset to MoveBack and Forth *****/
    Select @ScriptMove = 'Public Sub RecNext' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(05) + 'If not RS.EOF THEN' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'rs.movenext' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'loadrec' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(05) + 'End if' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'End Sub' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'Public Sub RecPrev' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'If not RS.BOF THEN' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'rs.movePrevious' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'loadrec' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'End if' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'End Sub' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'Public Sub RecFirst' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'If not RS.BOF THEN' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'rs.moveFirst' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'loadrec' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'End if' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'End Sub' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'Public Sub RecLast' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'If not RS.EOF THEN' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'rs.moveLast' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(10) + 'loadrec' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'End if' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'End Sub' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'Public Function RecCount() as long' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + space(5) + 'reccount = rs.recordcount' + char(13) + char(10)
    Select @ScriptMove = @ScriptMove + 'End Function' + char(13) + char(10)

	SELECT @ClassValidation = @ClassValidation + 'End Function' + char(13) + char(10)
	Select @CriteraSQL = 'Public Property Let Criteria(sNewCriteria as string)'+ char(13) + char(10) +
		space(5) + 'sCriteria = sNewCriteria' + char(13) + char(10) + 'End Property'
	Select @CriteraSQL = @CriteraSQL + 'Public Property Get Criteria() as string'+ char(13) + char(10) +
		space(5) + 'Criteria = sCriteria' + char(13) + char(10) + 'End Property'
	Select @CriteraSQL = @CriteraSQL + 'Public sub CriteriaReset()' + char(13) + char(10) +
		space(5) + 'sCriteria = ""' + char(13) + char(10) + 'End Sub'
    /*****  Print the Class Here *****/
    Print '''Class Start'
    Print @ClassStart
    print char(13) + char(10)
    Print @ClassType  
    print char(13) + char(10)
    Print @SQL5
    print char(13) + char(10)

	Print @LoadClass + 	
      Space(05) + 'If RS.State = adStateOpen Then RS.Close'+ char(13) + char(10) +
      space(05) + 'sSQL = "SELECT ' + 
      Left(@SQLSelect, Len(@SQLSelect) -1) + '" _' + char(13) + char(10) + space(10) + 
      '& "' + ' From ' + @TableName + '"' + char(13) + char(10)
	Print @LoadClass3	
    Print space(5) + 'RS.Open ssql, cnCLS' + char(13) + char(10)
    Print space(5) + 'If not ForArray Then LoadRec'+ char(13) + char(10)
    Print Space(5) + 'ForArray = False'
    Print 'End Sub'
    print char(13) + char(10)

    print @LoadRec + Space(10) + 'IsNewRec = False' + char(13) + char(10) + 
        @LoadRec1 + Space(10) + 'IsNewRec = True' + char(13) + char(10) + 
        Space(5) +  'End IF' + char(13) + char(10) + char(9) + 'Orig = Rec' 
    print char(13) + char(10) + 'End Sub'
    print char(13) + char(10)
	Print @ClassValidation	
	print char(13) + char(10)
	
    print @ScriptMove
    print char(13) + char(10)
    print @SQLArray
    print char(13) + char(10)
    print @SQLSave1
    print char(13) + char(10)
    print @SQLSave2
    	print '''*****  Property Calls' + char(13) + char(10)
	Print 'Public Property Get InvalidFields() As Variant' + char(13) + char(10) + Space(5) +
		'InvalidFields = InValid' + char(13) + char(10) + 
		'End Property' + char(13) + char(10) + char(13) + char(10)

    print @PropertyAll
    Print '''Class Init and Term'
    Print @SQL6
  END