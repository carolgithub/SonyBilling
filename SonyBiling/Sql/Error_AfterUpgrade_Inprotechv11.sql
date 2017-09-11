select  distinct(Case when Isnull(O.Currency,'') = '' then 'SGD' else O.currency end) as Currency, O.ITEMDATE as InvDate, O.OpenitemNo as InvNo, 
C.IRN,  
--wh1.MOVEMENTCLASS , WH1.REFTRANSNO , wh3.reftransno, wh1.transno, wh3.transno,
(Select distinct Reports.dbo.fnFileRef_LF(officialNumber) from ECSF.DBO.OfficialNumbers R where CaseID = C.CaseID and NumberType = 'Z') as FileRef, 
(Select distinct officialNumber from ECSF.DBO.OfficialNumbers R where CaseID = C.CaseID and NumberType = 'Z')  as FileRef_SF, 
(Select distinct AbbreviatedName from ECSF.DBO.Employee where Employeeno = E.EmployeeNo) as EmpCode , 
(Select Top 1 WX.WIPcode from ECSF.DBO.WORKHISTORY WX where 
	  WX.REFTRANSNO = b.ITEMTRANSNO AND WX.BILLLINENO = b.ITEMLINENO and wX.MOVEMENTCLASS = 2 and wX.status <> 0) as ChargeCode, 
--B.WIPCODE as ChargeCode,  --- 4 Sep UPD to captute WIPcode from Workhistory
(SELECT distinct Narrativecode from ECSF.DBO.NARRATIVE where NARRATIVENO = B.NarrativeNo) as NarrativeCode,  
(Case when ltrim(rtrim(IsNull(b.shortnarrative,''))) ='' then ltrim(rtrim(Convert(varchar(8000),na.NarrativeText))) else ltrim(rtrim(IsNull(b.shortnarrative,''))) end ) as NarrativeText, 
B.FOREIGNVALUE as Amount,Null as doctype   from ECSF.DBO.OPENITEM O  
join ECSF.DBO.BILLLINE B on (B.ITEMENTITYNO=O.ITEMENTITYNO  and B.ITEMTRANSNO =O.ITEMTRANSNO) 
join ECSF.DBO.CASES C on (C.IRN=B.IRN) 
join ECSF.DBO.WORKHISTORY wh2 on (b.ITEMTRANSNO = wh2.reftransno  and b.ITEMLINENO = wh2.billlineno)  
join ECSF.DBO.WORKHISTORY wh1 ON (WH1.REFTRANSNO = b.ITEMTRANSNO AND WH1.BILLLINENO = b.ITEMLINENO )  
LEFT JOIN ECSF.DBO.WORKHISTORY WH3 ON (WH3.REFTRANSNO = WH1.REFTRANSNO AND WH3.TRANSNO = WH1.TRANSNO  AND WH3.DISCOUNTFLAG = 1 )  
LEFT JOIN ECSF.DBO.BILLLINE b1 on (b1.ITEMTRANSNO = wh3.reftransno AND b1.ITEMLINENO = WH3.BILLLINENO)  
left join ECSF.DBO.CASENAME EMP on (EMP.CASEID=C.CASEID and EMP.NAMETYPE='ATT' and EMP.EXPIRYDATE is null) left join ECSF.DBO.WIPCATEGORY W on (W.CATEGORYCODE=B.CATEGORYCODE) 
left join ECSF.DBO.EMPLOYEE E   on (E.EMPLOYEENO=isnull((select min(WH.EMPLOYEENO) 
														from ECSF.DBO.WORKHISTORY WH Where WH.REFENTITYNO=O.ITEMENTITYNO and WH.REFTRANSNO =O.ITEMTRANSNO and WH.BILLLINENO =B.ITEMLINENO and WH.EMPLOYEENO is not null),EMP.NAMENO))  
left join ECSF.DBO.TABLECODES T on (T.TABLECODE=E.STAFFCLASS) 
left join ECSF.DBO.ASSOCIATEDNAME AN on (AN.NAMENO=O.ACCTDEBTORNO and AN.RELATIONSHIP='RES' and AN.JOBROLE=101347) 
left join ECSF.DBO.NAME N   on (N.NAMENO=AN.RELATEDNAME)  
left join ECSF.DBO.COUNTRY CN   on (CN.COUNTRYCODE=C.COUNTRYCODE) left join ECSF.DBO.CASETYPE CT  on (CT.CASETYPE=C.CASETYPE)  
left join ECSF.DBO.VALIDPROPERTY VP on (VP.PROPERTYTYPE=C.PROPERTYTYPE and VP.COUNTRYCODE=(Select min(VP1.COUNTRYCODE) from ECSF.DBO.VALIDPROPERTY VP1 where VP1.PROPERTYTYPE=VP.PROPERTYTYPE and VP1.COUNTRYCODE in (C.COUNTRYCODE, 'ZZZ')))  
left join ECSF.DBO.CASENAME REF on (REF.CASEID=C.CASEID and REF.NAMETYPE=CASE WHEN(O.RENEWALDEBTORFLAG=1) THEN 'Z' ELSE 'D' END and REF.NAMENO=O.ACCTDEBTORNO and REF.EXPIRYDATE is null)  
LEFT JOIN ECSF.DBO.FEESCALCULATION FC ON (O.ACCTDEBTORNO = FC.DEBTOR AND B.WIPCODE = FC.DISBWIPCODE AND B.NARRATIVENO = FC.DISBNARRATIVE  AND FC.DISBWIPCODE = 'D0009' AND FC.DISBNARRATIVE = 712 )  
LEFT JOIN ECSF.DBO.OFFICIALNUMBERS ONU ON (ONU.CASEID = C.CASEID AND ONU.NUMBERTYPE = 'A') LEFT JOIN ECSF.DBO.NARRATIVE NA ON (NA.NARRATIVENO = b.NARRATIVENO) 
LEFT JOIN ECSF.DBO.CASETEXT CX ON (CX.CASEID = C.CASEID AND CX.TEXTTYPE = 'BR') 
Where O.ACCTDEBTORNO in (-56054,-22869) and o.Status = 1 
and wh2.MOVEMENTCLASS = 2 
--and wh1.status <> 0 and wh2.status <> 0 and wh3.status <> 0
--and o.Openitemno = '1726941'
and O.ITEMDATE Between '08/01/2017 00:00:00' And  '08/31/2017 23:59:59'

--select * from ecsf.dbo.Billline where ItemTransNo = 1869042
--select * from 