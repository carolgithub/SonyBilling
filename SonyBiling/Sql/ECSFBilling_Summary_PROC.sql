USE [Reports]
GO
/****** Object:  StoredProcedure [dbo].[ECSFBilling_Summary_PROC]    Script Date: 11/09/2017 9:02:28 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****************Script date 31/05/04  5 PM**********************************************/

ALTER PROCEDURE [dbo].[ECSFBilling_Summary_PROC]
@FromDate varchar(20), @Todate varchar(20) as

declare  @InvoiceNo_t varchar(30), @InvoiceNo_d varchar(30), @BillCode_t varchar(5),
@SFee_t numeric(9,2), @OffFee_t numeric(9,2), @OthFee_t numeric(9,2), @TotFee_t numeric(9,2),
@cnt_t int, @InitVal int, @Init int, @dbn_t varchar(30)

--Drop Procedure ECSFBilling_Summary_PROC
--Exec ECSFBilling_Summary_PROC '2017/08/01 00:00:00',  '2017/08/31 23:59:59'

/*

select totalamount, billcode from BillTransactionDetails
where chargeddate between '2012/07/01 00:00:00' and '2012/06/30 23:59:59'
order by Invoiceno

*/

SET @FromDate = @FromDate + ' 00:00:00'
SET @Todate = @Todate + ' 23:59:59'

print @FromDate
print @Todate

/*set dateformat dmy  'ISG06256/04' '06/01/2004' and '06/30/2004' */

if exists (select object_id ( 'BillTransactionSummaryRep' ))
delete from BillTransactionSummaryRep

declare Cur_DisInv scroll cursor for
select distinct(Invoiceno) from BillTransactionDetails
where chargeddate between @FromDate and @Todate
order by Invoiceno

/*
select * from BillTransactionDetails where chargeddate between '05/01/2012 00:00:00' and '05/31/2012 23:59:59'
*/

open Cur_DisInv
fetch first from Cur_DisInv into @InvoiceNo_d

print @InvoiceNo_d

while @@fetch_status = 0 
begin	
	declare Cur_Sum scroll cursor for
	select BillTransactionDetails.BillCode, Invoiceno, sum(ServiceFee),
	sum(OfficialFee), sum(Others), sum(TotalAmount), DebitNoteNo
	from BillTransactionDetails, BillCode
	where BillCode.BillCode = BillTransactionDetails.BillCode	
	and invoiceno = @InvoiceNo_d
	group by BillTransactionDetails.BillCode, Description, Invoiceno, DebitNoteNo
	order by BillTransactionDetails.BillCode, Invoiceno, DebitNoteNo
	
	open Cur_Sum
	fetch first from Cur_Sum into 
	@BillCode_t , @InvoiceNo_t , @SFee_t , @OffFee_t , @OthFee_t, @TotFee_t, @dbn_t

	Select @InitVal = 0
	Select @Init = 1

	while @@fetch_status = 0 
	begin	
		
		/*Check if BillCode = 5 by itself
		Else change the Billode from 5 to other BillCodes(1,2,3,4) */ 
		print 'hell2'
		if (select @Init) = 1 
		Begin
			Select @Init = 0
			Select @InitVal = @BillCode_t
			print @Initval
		End
			Print @BillCode_t
			if exists(select Billcode from BillTransactionDetails 
			          where Invoiceno = @InvoiceNo_t and Billcode <> '5')
			Begin
				if  (select @BillCode_t) = '5' 
				Begin
					insert into BillTransactionSummaryRep values (@InvoiceNo_t , 
					@dbn_t, @Initval , @SFee_t, @OffFee_t, @OthFee_t, @TotFee_t)
				end
				else 
				Begin	
					insert into BillTransactionSummaryRep values (@InvoiceNo_t , 
					@dbn_t, @BillCode_t , @SFee_t, @OffFee_t, @OthFee_t, @TotFee_t)
				end
			end
			else
			Begin				
				if  (select @BillCode_t) = '5' 
				Begin
					insert into BillTransactionSummaryRep values (@InvoiceNo_t , 
					@dbn_t, @BillCode_t , @SFee_t, @OffFee_t, @OthFee_t, @TotFee_t)
				end
			End		

		fetch next from Cur_Sum into 
		@BillCode_t , @InvoiceNo_t , @SFee_t , @OffFee_t , @OthFee_t, @TotFee_t, @dbn_t

		
	end
	close Cur_Sum
	deallocate Cur_Sum

fetch next from Cur_DisInv into @InvoiceNo_d

end

close Cur_DisInv
deallocate Cur_DisInv

