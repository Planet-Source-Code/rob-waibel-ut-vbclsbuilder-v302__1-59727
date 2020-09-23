SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


/*	Procedure for 8.0 server */
CREATE PROCEDURE sp_fkeysRDW(
			   @pktable_name		sysname = null,
			   @fktable_name		sysname = null)
as
	set nocount on
	DECLARE @pktable_id			int
	DECLARE @pkfull_table_name	nvarchar(257) /* 2*128 + 1 */
	DECLARE @fktable_id			int
	DECLARE @fkfull_table_name	nvarchar(257) /* 2*128 + 1 */
	declare	@order_by_pk		int
	Declare 	@@pktable_name		sysname ,
			@@PKCOLUMN_NAME	sysname ,
			@@fktable_name		sysname ,
			@@FKCOLUMN_NAME	sysname 

	/* select 'XXX starting table creation' */

	create table #fkeysall(
			rkeyid int NOT NULL,
			rkey1 int NOT NULL,
				rkey2 int NOT NULL,
				rkey3 int NOT NULL,
				rkey4 int NOT NULL,
				rkey5 int NOT NULL,
				rkey6 int NOT NULL,
				rkey7 int NOT NULL,
				rkey8 int NOT NULL,
				rkey9 int NOT NULL,
				rkey10 int NOT NULL,
				rkey11 int NOT NULL,
				rkey12 int NOT NULL,
				rkey13 int NOT NULL,
				rkey14 int NOT NULL,
				rkey15 int NOT NULL,
				rkey16 int NOT NULL,
			fkeyid int NOT NULL,
			fkey1 int NOT NULL,
				fkey2 int NOT NULL,
				fkey3 int NOT NULL,
				fkey4 int NOT NULL,
				fkey5 int NOT NULL,
				fkey6 int NOT NULL,
				fkey7 int NOT NULL,
				fkey8 int NOT NULL,
				fkey9 int NOT NULL,
				fkey10 int NOT NULL,
				fkey11 int NOT NULL,
				fkey12 int NOT NULL,
				fkey13 int NOT NULL,
				fkey14 int NOT NULL,
				fkey15 int NOT NULL,
				fkey16 int NOT NULL,
			constid int NOT NULL,
			name sysname collate database_default NOT NULL)

	create table #fkeys(
			pktable_id		int NOT NULL,
			pkcolid 		int NOT NULL,
			fktable_id		int NOT NULL,
			fkcolid 		int NOT NULL,
			KEY_SEQ 		smallint NOT NULL,
			fk_id			int NOT NULL,
			PK_NAME			sysname collate database_default NOT NULL)

	create table #fkeysout(
			PKTABLE_QUALIFIER sysname collate database_default NULL,
			PKTABLE_OWNER sysname collate database_default NULL,
			PKTABLE_NAME sysname collate database_default NOT NULL,
			PKCOLUMN_NAME sysname collate database_default NOT NULL,
			FKTABLE_QUALIFIER sysname collate database_default NULL,
			FKTABLE_OWNER sysname collate database_default NULL,
			FKTABLE_NAME sysname collate database_default NOT NULL,
			FKCOLUMN_NAME sysname collate database_default NOT NULL,
			KEY_SEQ smallint NOT NULL,
			UPDATE_RULE smallint NULL,
			DELETE_RULE smallint NULL,
			FK_NAME sysname collate database_default NULL,
			PK_NAME sysname collate database_default NULL,
			DEFERRABILITY smallint null)

	/* select 'XXX starting parameter analysis' */

    select  @order_by_pk = 0

	if (@pktable_name is null) and (@fktable_name is null)
	begin	/* If neither primary key nor foreign key table names given */
		raiserror (15252,-1,-1)
		return
    end

	/*	Get Object ID */
	if @pktable_name is not null BEGIN
		SELECT @pkfull_table_name = quotename(@pktable_name)
		SELECT @pktable_id = object_id(@pkfull_table_name)
		END
	/*	Get Object ID */
	if @fktable_name is not null BEGIN
		SELECT @fkfull_table_name = quotename(@fktable_name)
		SELECT @fktable_id = object_id(@fkfull_table_name)
		END
	if @fktable_name is not null
	begin
		if @fktable_id is null
			SELECT @fktable_id = 0	/* fk table not found, empty result */
    end

	if @pktable_name is null
	begin /*  If table name not supplied, match all */
		select @order_by_pk = 1
	end
	else
	begin
		if @pktable_id is null
		begin
			SELECT @pktable_id = 0	/* pk table not found, empty result */
		end
	end

	/*	SQL Server supports upto 16 PK/FK relationships between 2 tables */
	/*	Process syskeys for each relationship */
	/*  First, attempt to get all 16 keys for each rel'ship, then sort
		them out with a 16-way "insert select ... union select ..." */

	/* select 'XXX starting data analysis' */

	insert into #fkeysall
		select
			r.rkeyid,
			r.rkey1, r.rkey2, r.rkey3, r.rkey4,
				r.rkey5, r.rkey6, r.rkey7, r.rkey8,
				r.rkey9, r.rkey10, r.rkey11, r.rkey12,
				r.rkey13, r.rkey14, r.rkey15, r.rkey16,
			r.fkeyid,
			r.fkey1, r.fkey2, r.fkey3, r.fkey4,
				r.fkey5, r.fkey6, r.fkey7, r.fkey8,
				r.fkey9, r.fkey10, r.fkey11, r.fkey12,
				r.fkey13, r.fkey14, r.fkey15, r.fkey16,
			r.constid,
			i.name
		from
			sysreferences r, sysobjects o, sysindexes i
		where	r.constid = o.id
			AND o.xtype = 'F'
			AND r.rkeyindid = i.indid
			AND r.rkeyid = i.id
			AND r.rkeyid between isnull(@pktable_id, 0)
							and isnull(@pktable_id, 0x7fffffff)
			AND r.fkeyid between isnull(@fktable_id, 0)
							and isnull(@fktable_id, 0x7fffffff)

	/* select count (*) as 'XXX countall' from #fkeysall */

    insert into #fkeys
			select rkeyid, rkey1, fkeyid, fkey1, 1, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey2, fkeyid, fkey2, 2, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey3, fkeyid, fkey3, 3, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey4, fkeyid, fkey4, 4, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey5, fkeyid, fkey5, 5, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey6, fkeyid, fkey6, 6, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey7, fkeyid, fkey7, 7, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey8, fkeyid, fkey8, 8, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey9, fkeyid, fkey9, 9, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey10, fkeyid, fkey10, 10, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey11, fkeyid, fkey11, 11, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey12, fkeyid, fkey12, 12, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey13, fkeyid, fkey13, 13, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey14, fkeyid, fkey14, 14, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey15, fkeyid, fkey15, 15, constid, name
			from #fkeysall
		union all
			select rkeyid, rkey16, fkeyid, fkey16, 16, constid, name
			from #fkeysall

	/* select count (*) as 'XXX count' from #fkeys */

	insert into #fkeysout
		select
			PKTABLE_QUALIFIER = convert(sysname,db_name()),
			PKTABLE_OWNER = convert(sysname,USER_NAME(o1.uid)),
			PKTABLE_NAME = convert(sysname,o1.name),
			PKCOLUMN_NAME = convert(sysname,c1.name),
			FKTABLE_QUALIFIER = convert(sysname,db_name()),
			FKTABLE_OWNER = convert(sysname,USER_NAME(o2.uid)),
			FKTABLE_NAME = convert(sysname,o2.name),
			FKCOLUMN_NAME = convert(sysname,c2.name),
			KEY_SEQ,
			UPDATE_RULE = CASE WHEN (ObjectProperty(fk_id, 'CnstIsUpdateCascade')=1) THEN 
				convert(smallint,0) ELSE convert(smallint,1) END,
			DELETE_RULE = CASE WHEN (ObjectProperty(fk_id, 'CnstIsDeleteCascade')=1) THEN 
				convert(smallint,0) ELSE convert(smallint,1) END,
			FK_NAME = convert(sysname,OBJECT_NAME(fk_id)),
			PK_NAME,
			DEFERRABILITY = 7	/* SQL_NOT_DEFERRABLE */
		from #fkeys f,
			sysobjects o1, sysobjects o2,
			syscolumns c1, syscolumns c2
		where	o1.id = f.pktable_id
			AND o2.id = f.fktable_id
			AND c1.id = f.pktable_id
			AND c2.id = f.fktable_id
			AND c1.colid = f.pkcolid
			AND c2.colid = f.fkcolid
	/* select count (*) as 'XXX countout' from #fkeysout */

	Declare FK Cursor For 
		Select	PKTABLE_NAME, 
			  	PKCOLUMN_NAME, 
			  	FKTABLE_NAME,
			  	FKCOLUMN_NAME from #fkeysout
		order by 1,2,3
	For Read Only
	Open FK
	Fetch Next From FK into 	@@pktable_name ,
						@@PKCOLUMN_NAME ,
						@@fktable_name ,
						@@FKCOLUMN_NAME
	WHile @@Fetch_Status = 0 BEGIN
		Print '''References Table ' + @@pktable_name + ' on field ' + @@PKCOLUMN_NAME +
			' to Table ' + @@fktable_name + ' on field ' + @@FKCOLUMN_NAME
		Fetch Next From FK into 	@@pktable_name , @@PKCOLUMN_NAME , @@fktable_name , @@FKCOLUMN_NAME
		END
	Deallocate FK

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

