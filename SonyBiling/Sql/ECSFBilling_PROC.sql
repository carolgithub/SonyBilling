USE [Reports]
GO
/****** Object:  StoredProcedure [dbo].[ECSFBilling_PROC]    Script Date: 11/09/2017 8:48:24 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************Script date 31/05/04  5 PM Updated on 03/03/2005**********************************************/
ALTER PROCEDURE [dbo].[ECSFBilling_PROC] 
@ClientCode_var varchar(10), @DebitNoteNo_var varchar(20) , @MEndDate_var Datetime as 

--exec ECSFBilling_PROC '01481', 'DBN08/17',  '2017/08/01'
--FrDate Datetime, @ToDate Datetime as

declare @CurrencyCode_t varchar(10) , @ChargedDate_t datetime , 
@InvoiceNo_t varchar(20), @FileRefNo_t varchar(12) , @FileRefNo_SF_t varchar(12) , @InvGST varchar(20),
@AttorneyCode_t varchar(10) , @BillingCode_t varchar(10),  @Narrative_t varchar(10),@TotalAmount_t numeric(9,2), 
@HrRate_var varchar(10), @BillCode_var varchar(2), @ClientBillingCode_var varchar(10), 
@AppNo_var varchar(50), @PatentNo_var varchar(50), @AttorneyName_var varchar(30),@WIPCode_t varchar(10), 
@ClientRefNo_var varchar(15), @ClCode_var varchar(3), @DocType_var varchar(1), @WIPType nvarchar(10)
 
Delete from BillTransactionDetails where DebitNoteNo = @DebitNoteNo_var 

--Delete from BillTransactionDetails where DebitNoteNo = 'DBN08/17'
--select * from BillTransactionDetails where DebitNoteNo like 'DBN16/02%'/16'
--select * from BillTransactionDetails where InvoiceNo = '1300174'
--select @ClientCode_var = (SELECT REPLACE(@CurrencyCode_t,'$','D'))  ('01481', 'SONY1')

--select * from BillTransactionDetails where filerefno in ('1481SG1673', '1481SG1674' , '1481SG1672' )
--select * from BillTransactionDetails where filerefno in ('36085sg00003', '36085sg00004' , '36085sg00005' )
--delete from BillTransactionDetails where filerefno in ('36085sg00003', '36085sg00004' , '36085sg00005' )
--select * from BillTransactionDetails where filerefno like '36085SG%'

--36085sg3, 36085sg4, 36085sg5 (transferred to Velos Media )  --(1481SG1673, 1481SG1674 , 1481SG1672 )


declare Cur_Trans scroll cursor for
Select CurrencyCode , ChargedDate , InvoiceNo, FileRefNo , FileRefNo_sf , AttorneyCode , 
Isnull(WIPCode,''), (Select WIPTypeID from ecsf.dbo.WIPTemplate where WIPCODE = Isnull(t.WIPCode,'')) as WIPType,
IsNull(NarrativeCode,''),TotalAmount, DocType from BillTransactionImport t
order by InvoiceNo 

open Cur_Trans
fetch first from Cur_Trans into 
@CurrencyCode_t , @ChargedDate_t , @InvoiceNo_t , @FileRefNo_t , @FileRefNo_SF_t ,@AttorneyCode_t ,  
@WIPCode_t ,@WIPTYpe, @Narrative_t, @TotalAmount_t , @DocType_var

while @@fetch_status = 0 
begin	

			/*Select CurrencyCode*/
			if exists(select CHARINDEX('$', @CurrencyCode_t, 1))

			begin
				select @CurrencyCode_t = (SELECT REPLACE(@CurrencyCode_t,'$','D'))
			end		

			----------------------------------------------------------------------------------
			if exists(select ClientBillingCode from BillingCode where  WIPCode = @WIPCode_t)			  
			begin			
			
				insert into BillTransactionDetails 
						
				select 'SG' , '03SG' , @CurrencyCode_t , 
				@ChargedDate_t , @MEndDate_var, @DebitNoteNo_var , @InvoiceNo_t , @ClientCode_var, 	
				Null,  @FileRefNo_t , 					
				(select AttorneyName from AttorneyFee where AttorneyCode = @AttorneyCode_t),
				(@CurrencyCode_t + (select HourlyFee from AttorneyFee where AttorneyCode = @AttorneyCode_t)),
				(select distinct(isnull(BillCode,'')) from BillingCode where Isnull(WIPCode,'') = @WIPCode_t and Isnull(NarrativeCode,'') = @Narrative_t) , 
				Null As ApplicationNo,Null as PatentNo,	
				(select Case when isnull(ClientBillingCode,'') = '' then '' else ClientBillingCode end from BillingCode where Isnull(WIPCode,'') = @WIPCode_t and Isnull(NarrativeCode,'') = @Narrative_t), 
				@WIPCode_t , @Narrative_t, 
				(select Case when (substring(@WIPType,1,3)  = 'SER' OR @WIPType like '%LOD%' ) then @TotalAmount_t else '0.00' end),
				(select Case when charindex('FEE', @WIPType,1)  <> 0 then @TotalAmount_t else '0.00' end),
				(select Case when (substring(@WIPType,1,3) <> 'SER' and @WIPType NOT like '%LOD%' and charindex('FEE', @WIPType,1)  = 0 ) then @TotalAmount_t else '0.00' end),
				@TotalAmount_t , Null, @DocType_var 

			end			

Select @ClientBillingCode_var = ''
Select @BillCode_var = ''
Select @HrRate_var = ''
Select @AppNo_var = ''
Select @PatentNo_var = ''
select @AttorneyName_var = ''
select @ClientRefNo_var = ''
select @ClCode_var = ''
select @DocType_var = ''

fetch next from Cur_Trans into 
@CurrencyCode_t , @ChargedDate_t , @InvoiceNo_t , @FileRefNo_t , @FileRefNo_SF_t ,@AttorneyCode_t , 
@WIPCode_t ,@WIPTYpe, @Narrative_t, @TotalAmount_t , @DocType_var

end

close Cur_Trans
deallocate Cur_Trans


------------------------------------------------------------------------------------------------
	--------1 Apr 2011------------------------------------------------------------------------------

		/*Update BillTransactionDetails set ClientRefNo = (select distinct(Dept_ref) from DBServer.Diams.dbo.Pat p, 
		DBServer.Diams.dbo.pat_Depts t, DBServer.Diams.dbo.L_Alternate_Key l, DBServer.Diams.dbo.pat_date d,
		BillTransactionDetails z where p.patpk = t.patfk and l.record_fk = d.record_fk 
		and p.patpk = d.record_fk and (ALTERNATE_KEY = FileRefNo or ALTERNATE_KEY = dbo.fnFileRef_SF(FileRefNo)) 
		and (Doc_No = FileRefNo or Doc_No = dbo.fnFileRef_SF(FileRefNo))   and DebitNoteNo = @DebitNoteNo_var)*/

		Update d set ClientRefNo = ReferenceNo from ecsf.dbo.CASENAME m, ecsf.dbo.Cases c ,
		BillTransactionDetails d, BillTransactionImport p where c.CASEID = m.CASEID and NAMETYPE = 'D' and c.IRN = p.IRN
		and d.InvoiceNo = p.InvoiceNo and DebitNoteNo =  @DebitNoteNo_var
	
		--To update ClientRefNo if its wrong-----------------
		--Select substring('S11P0253SG00 (SP253817SG00)', Charindex('(', 'S11P0253SG00 (SP253817SG00)') + 1, (Charindex(')','S11P0253SG00 (SP253817SG00)')-Charindex('(','S11P0253SG00 (SP253817SG00)')) -1 )
		Update d set ClientRefNo = substring(ClientRefNo, Charindex('(', ClientRefNo) + 1, (Charindex(')',ClientRefNo)-Charindex('(',ClientRefNo)) -1 )
		from BillTransactionDetails d where DebitNoteNo =  @DebitNoteNo_var
		and ISNULL(ClientRefNo,'') <> '' and  Charindex('(', ClientRefNo) <> 0
		
		Update d set ApplicationNo = OFFICIALNUMBER
		from ecsf.dbo.OFFICIALNUMBERS O, ecsf.dbo.CASES C , BillTransactionDetails d, BillTransactionImport p 
		WHERE C.CaseID = O.CaseID and NumberType = 'A' and ISCURRENT = 1 and C.IRN = p.IRN
		and d.InvoiceNo = p.InvoiceNo and DebitNoteNo =  @DebitNoteNo_var 

		Update d set PatentNo = OFFICIALNUMBER
		from ecsf.dbo.OFFICIALNUMBERS O, ecsf.dbo.CASES C , BillTransactionDetails d, BillTransactionImport p 
		WHERE C.CaseID = O.CaseID and NumberType = 'R' and ISCURRENT = 1 and C.IRN = p.IRN
		and d.InvoiceNo = p.InvoiceNo and DebitNoteNo =  @DebitNoteNo_var 

			/*To Check the ClientRefNo = 12 digits*/
			if exists(select  distinct(FileRefNo) from BillTransactionDetails where ( len(ClientRefNo) < 12 or 
				 len(ClientRefNo) > 12 or isnull(ClientRefNo, '' ) = '' ) and DebitNoteNo =  @DebitNoteNo_var)
				 			  
			begin	
				insert into BillTransactionError select 'Invalid ClientReference No', ChargedDate, InvoiceNo, FileRefNo, AttorneyName, WIPCode, NarrativeCode,TotalAmount
				from BillTransactionDetails where ( len(ClientRefNo) < 12 or 
				len(ClientRefNo) > 12 or isnull(ClientRefNo, '' ) = '' ) and DebitNoteNo =  @DebitNoteNo_var  
	
			
			End
		
			/*To Check if AppNo and PatNo missing*/
			if exists(select distinct(FileRefNo) from BillTransactionDetails where isnull(ApplicationNo, '' ) = ''  
				  and isnull(PatentNo, '' ) = '' and DebitNoteNo =  @DebitNoteNo_var  )			  
			begin	
				insert into BillTransactionError select 'Missing AppNo and PatNo', ChargedDate, InvoiceNo, FileRefNo, AttorneyName, WIPCode, NarrativeCode,TotalAmount
				from BillTransactionDetails where isnull(ApplicationNo, '' ) = ''  
				 and isnull(PatentNo, '' ) = '' and DebitNoteNo =  @DebitNoteNo_var  	
			End
		
			
			/*To Check Invalid Billing Codes(Updated)*/
		
			If Not exists(select distinct(b.WIPCode) from BillTransactionImport b, BillingCode C where 
			IsNull(b.WIPCode,'' ) = ISNULL(c.WIPCode,'') and IsNull(b.NarrativeCode,'' ) = ISNULL(c.NarrativeCode,'') )
		  
			begin	
				insert into BillTransactionError select 'Invalid Billing Code', t.ChargedDate, t.InvoiceNo, t.FileRefNo, 
				t.AttorneyCode, t.WIPCode, t.NarrativeCode, t.TotalAmount
				from BillTransactionImport t , BillingCode C where 
			    IsNull(t.WIPCode,'' ) = ISNULL(c.WIPCode,'') and IsNull(t.NarrativeCode,'' ) = ISNULL(c.NarrativeCode,'')
 		
			End
		
			/*To Check missing Atorney*/
			if exists(select distinct(FileRefNo) from BillTransactionDetails where isnull(AttorneyName, '' ) = ''  
				  and DebitNoteNo =  @DebitNoteNo_var )			  
			begin	
				insert into BillTransactionError select 'Missing Attorney', ChargedDate, InvoiceNo, FileRefNo, AttorneyName, WIPCode, NarrativeCode,TotalAmount
				from BillTransactionDetails where isnull(ApplicationNo, '' ) = ''  
				and isnull(PatentNo, '' ) = '' and DebitNoteNo =  @DebitNoteNo_var  	
			End

	------------------------------------------------------------------------------------------------
Update t set AttorneyName = 'R.N.GNANAPRAGASAM' from  BillTransactionDetails t 
where DebitNoteNo =@DebitNoteNo_var 

	-------------------------------------------------------------------------------------------------
/**********************************************************/

if exists (select object_id ( 'BillTransactionError' ))
delete from BillTransactionError



/*

--To Delete--------------------------
Select * from BillTransactionImport where filerefno in ('34132SG00005', '34132SG00007')
Delete from  BillTransactionImport where filerefno_SF in ('1481SG1623', '1481SG1639')

Update B set FileRefNo = '01481SG01710',
			 FileRefNo_SF = '1481SG1710'  from BillTransactionImport B where FileRefNo = '34132SG00005'
			 
Update B set FileRefNo = '01481SG01714',
			 FileRefNo_SF = '1481SG1714'  from BillTransactionImport B where FileRefNo = '34132SG00007'

Select * from BillTransactionImport where InvoiceNo in ('1210915CN', '1210916CN' )
Delete from  BillTransactionImport where InvoiceNo in ('1207788CN', '1206951CN')
Delete from  BillTransactionImport where InvoiceNo in ('1502098CN')
----------------------------------------------------------
select sum(Totalamount) from BillTransactionImport 
select sum(Totalamount) from BillTransactionDetails where DebitNoteNo = 'DBN08/17'


---UPDATE -----IMP----------------------------------------
select sum(Totalamount) from BillTransactionImport 
select sum(Totalamount) from BillTransactionDetails where DebitNoteNo = 'DBN08/17'
Select * from BillTransactionImport where InvoiceNo = '1317241'

select * from BillingCode  order by wipcode, Narrativecode
Select * from BillTransactionImport where NarrativeCode is Null order by narrativecode, wipcode

1.	Update t set NarrativeCode = 'SP0069', WIPcode = 'S0081'	
	--,WIPCode = 'D0002'
	from BillTransactionImport t where
	 --InvoiceNo = '1531363' and 
	WIPCode =  'S0075'	 and
	isnull(NarrativeCode,'') = ''
	--isnull(NarrativeCode,'') = 'DP0056'
	--and convert(varchar(1000),Narrative) = 'PTO fees for search/ examination/ search & examination PF11'
	
		Update t set WIPCode = 'D0013' ,
		 Narrativecode  = 'DP0030' 
		from BillTransactionImport t where 
		--convert(varchar(2000),Narrative) = 'Other prosecution activities not covered by the above categories'
		WIPCode = 'D0003' 
		and Narrativecode is null 
		--and Narrativecode  = 'DP0030' 
		
	select * from BillTransactionImport where wipcode not in (select distinct WIPCode from BillingCode ) order by wipcode

	select * from BillingCode  order by wipcode, Narrativecode	

 2.	Update t set NarrativeCode = 'SP0014' , WIPcode = 'S0032' from BillTransactionImport t where 
	WIPCode = 'S0048' and 
	narrativecode =  'SP0014'
	--and
	--Invoiceno = '1625537'
	-- and
	 --and TotalAmount = 475.00 and totalamount = 150.00
	convert(varchar(2000),Narrative) = 'Preparing and filing voluntary amendments to the specification and preparing necessary forms.'

Update t set NarrativeCode = 'DP0067' from BillTransactionImport t where WIPCode = 'D0009' 
Update t set WIPCode = 'D9003' from BillTransactionImport t where WIPCode = 'D0009' and NarrativeCode = 'DP0067' 
	select * from BillingCode  order by wipcode, Narrativecode
		
3.  select * from BillTransactionDetails where DebitNoteNo = 'DBN08/17' 
    and (Billcode is Null or Billingcode is Null) 
    order by WIPcode  
    and WIPCODE = 'S0007' and Narrativecode = 'SP0062'

    Update t set WIPcode = 'S0039', NarrativeCode = 'SP0176' from BillTransactionImport t where 
	----Invoiceno in ('1705599') 
	WIPcode = 'S0039' 
	--and convert(varchar(2000),Narrative) = 'Dealing with request for examination / search & examination'
	and Narrativecode =  'SP0185' 
