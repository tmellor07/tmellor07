/*  Case: 2180363
   Author: Dmsi Terri
   Read the 3 excel file to create the report
   Excel file 1 - is created from the dispatch view
   Excel File 2 - is a file that has the driver rates that is maint by Dugar
   Excel file 3 - is a file that contains the totals mileage for each driver, this file comes from one of their dispatch apps

*/ 
 
define access hf_driver, set driver = DispatchViewer:driver
define file hfMileage = access hf_driver, set driver =  "Daily Opex/Mileage Rate"

Define String GroupCode = if DispatchViewer:carrier_name = "Lincoln Logistics, LLC" then "0" else "1" 
 
Define Number POID = if DispatchViewer:tran_type = "PO" then Tran_id
define access po_header,
   set po_id    = POID , generic , one to one

Define Number SOID = if DispatchViewer:tran_type = "SO" then Tran_id
define access so_header,
    set so_id    = SOID , generic , one to one

Define String SystemID2Use = 
      if DispatchViewer:tran_type = "SO" then so_header:system_id else Po_header:system_id

define access dispatch_header,
  set system_id     = SystemID2Use ,
  set dispatch_id   = DispatchViewer:dispatch, 
   one to one,  null fill on failure

define access shipment_info,
  Set system_id    = dispatch_header:system_id,
  set tran_sysid   = dispatch_header:system_id,
  set tran_type    = "Dp",
  set tran_id      = dispatch_header:dispatch_id, one to many 

define access so_cost_packet, 
set TRAN_ID_SYSID = dispatch_header:system_id,
set TRAN_TYPE = DispatchViewer:tran_type,
set TRAN_ID = DispatchViewer:tran_id ,
set shipment_num=DispatchViewer:shipment, null fill on failure

Define Number Frt = 
     if groupcode = "0" then
       if tran_type = "SO" and so_cost_packet:tran_id_sysid <> Null then so_cost_packet:cost 
       else dispatchviewer:freight_rate 
    else 50.00
Define Number Exp = if groupcode = "0" then  0 else dispatchviewer:freight_rate  

define number CityRate     = if stop = 1 then val(shipment_info:scac_code)/decimalplaces=2
define number DriverAmt    = CityRate * hf_driver:pay_rate
Define number LincolnRev   = if groupcode = "0"  then Frt
Define number CommonCarRev = if groupcode <> "0" then Frt

Include "miles.inc"

where  DispatchViewer:transport_type <> ""

list/duplicates/noreporttotals
/xls="performance_rpt"
   DispatchViewer:created_by 
   driver 
   carrier_name 
   dispatch/nocomma
   transport_type
   stop 
   tran_id
   shipment
   tran_type 
   origin_name
   origin_city 
   origin_state 
   destination_name 
   destination_city 
   destination_state 
   destination_address_1 /name="destaddr"

   Frt     /name="Frt"  
   Exp     /heading="Expense" 

   CityRate                        /Decimalplaces=2/duplicates/heading="SCAC-City Rate"
   hf_driver:pay_rate
   CityRate * hf_driver:pay_rate   /decimalplaces=2/heading="Driver Settlement"/name="Driveramt"
   dispatch_header:total_distance  /heading="Mileage"
 
   Val(shipment_info:stcc)        /decimalplaces=2/duplicates /heading="STCC-Deadhead Rate"
   if load_tarp= 0 then "No" else "Yes"/heading="Tarp"



Sorted by 
   GroupCode/newline
   dispatchviewer:dispatch stop

End of GroupCode/nohead
   Total[Frt]/align=Frt/tag="SubTotal"
   Total[exp]/align=exp/tag="SubTotal"
   Total[CityRate]/align=CityRate/tag="SubTotal"
   total[DriverAmt]/align=DriverAmt/tag="SubTotal"
   
End of Report/nohead
""/newline=2
"Total Lincoln Revenue"/align=destaddr    total[LincolnRev]/align=Frt/newline=2
"Common Carrier Rev"/align=destaddr       total[CommonCarRev]/align=Frt/newline=2
"Deiver Settlement"/align=destaddr        total[DriverAmt]/align=Frt/newline=2
"Operating Expenses"/align=destaddr     MileTotal * hfMileage:pay_rate /align=frt
       Str(MileTotal) + " x " + Str(hfMileage:pay_rate) /align=CityRate /newline=2
"Gross Margin"/align=destaddr
total[LincolnRev] +   total[LincolnRev] - total[driveramt]/align=Frt/newline
