# All Inputs for the LLM in a Single File

This consolidated file bundles all input files and artifacts that are available in the project according to the prompt
“🧭 Entry Point for Follow-Up LLM.”

Purpose:
To provide a single text file that can be uploaded to the LLM in scenarios where file uploads are not possible.

---

## 1. Context Sources (LLM Input)

### FILENAME: excelcell.csv
CONTENT:

Sheet,Address,Formula,Value
Calculation,A1,Endowment Life Insurance Calculator,Endowment Life Insurance Calculator
Calculation,J1,VBA solution with caching commutation values,VBA solution with caching commutation values
Calculation,A3,Policy data,Policy data
Calculation,D3,Tariff data,Tariff data
Calculation,G3,Limits,Limits
Calculation,J3,Premium calculation,Premium calculation
Calculation,A4,x,x
Calculation,B4,40,40.0
Calculation,D4,InterestRate,InterestRate
Calculation,E4,0.0175,0.0175
Calculation,G4,MinAgeFlex,MinAgeFlex
Calculation,H4,60,60.0
Calculation,A5,Sex,Sex
Calculation,B5,M,M
Calculation,D5,MortalityTable,MortalityTable
Calculation,E5,DAV1994_T,DAV1994_T
Calculation,G5,MinTermFlex,MinTermFlex
Calculation,H5,5,5.0
Calculation,J5,NormGrossAnnualPrem,NormGrossAnnualPrem
Calculation,K5,"=(act_nGrAx(x,n,Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x,Sex,MortalityTable,InterestRate)+gamma1*Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)+gamma2*(Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)-Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)))/((1-beta1)*Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)-alpha*t)",0.04226001029372492
Calculation,A6,n,n
Calculation,B6,30,30.0
Calculation,D6,alpha,alpha
Calculation,E6,0.025,0.025
Calculation,J6,GrossAnnualPrem,GrossAnnualPrem
Calculation,K6,=SumInsured*K5,"""4226.001"""
Calculation,A7,t,t
Calculation,B7,20,20.0
Calculation,D7,beta1,beta1
Calculation,E7,0.025,0.025
Calculation,J7,GrossModalPrem,GrossModalPrem
Calculation,K7,=(1+ModalSurcharge)/PayFreq*(K6+k),"""371.8751"""
Calculation,A8,SumInsured,SumInsured
Calculation,B8,100000,"""100000"""
Calculation,D8,gamma1,gamma1
Calculation,E8,0.0008,0.0008
Calculation,A9,PayFreq,PayFreq
Calculation,B9,12,12.0
Calculation,D9,gamma2,gamma2
Calculation,E9,0.00125,0.00125
Calculation,J9,Pxt,Pxt
Calculation,K9,"=(act_nGrAx(x,n,Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x,Sex,MortalityTable,InterestRate)+t*alpha*NormGrossAnnualPrem)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)",0.040012169238708255
Calculation,D10,gamma3,gamma3
Calculation,E10,0.0025,0.0025
Calculation,D11,k,k
Calculation,E11,24,24.0
Calculation,D12,ModalSurcharge,ModalSurcharge
Calculation,E12,"=IF(PayFreq=2,2%,IF(PayFreq=4,3%,IF(PayFreq=12,5%,0)))",0.05
Calculation,A14,Progression values,Progression values
Calculation,A15,k,k
Calculation,B15,Axn,Axn
Calculation,C15,axn,axn
Calculation,D15,axt,axt
Calculation,E15,kVx_pp,kVx_pp
Calculation,F15,kDRx_pp,kDRx_pp
Calculation,G15,kVx_pu,kVx_pu
Calculation,H15,kVx_MRV,kVx_MRV
Calculation,I15,Flex. phase,Flex. phase
Calculation,J15,Surrender deduction,Surrender deduction
Calculation,K15,Surrender value,Surrender value
Calculation,L15,SumInsured_pu,SumInsured_pu
Calculation,A16,0,0.0
Calculation,B16,"=IF(A16<=n,act_nGrAx(x+$A16,MAX(0,n-$A16),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A16,Sex,MortalityTable,InterestRate),0)",0.63159227805218
Calculation,C16,"=Act_axn_k(x+$A16,MAX(0,n-$A16),Sex,MortalityTable,InterestRate,1)",21.420277547537385
Calculation,D16,"=Act_axn_k(x+$A16,MAX(0,t-$A16),Sex,MortalityTable,InterestRate,1)",16.31309413156213
Calculation,E16,"=B16-P_xt*D16+gamma2*(C16-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D16)",-0.021130005146862474
Calculation,F16,=SumInsured*E16,"""-2113.0005"""
Calculation,G16,=B16+gamma3*C16,0.6851429719210235
Calculation,H16,"=F16+alpha*t*GrossAnnualPrem*Act_axn_k(x+A16,MAX(5-A16,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I16,"=IF(AND(x+A16>=MinAgeFlex,A16>=n-MinTermFlex),1,0)",0.0
Calculation,J16,"=IF(OR(A16>n,I16),0,MIN(150,MAX(50,1%*(SumInsured-F16))))","""150"""
Calculation,K16,"=MAX(0,H16-J16)","""0"""
Calculation,L16,"=IFERROR(IF(A16>n,0,IF(A16<t,H16/G16,SumInsured)),0)","""0"""
Calculation,A17,1,1.0
Calculation,B17,"=IF(A17<=n,act_nGrAx(x+$A17,MAX(0,n-$A17),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A17,Sex,MortalityTable,InterestRate),0)",0.6417247337591211
Calculation,C17,"=Act_axn_k(x+$A17,MAX(0,n-$A17),Sex,MortalityTable,InterestRate,1)",20.8311476228625
Calculation,D17,"=Act_axn_k(x+$A17,MAX(0,t-$A17),Sex,MortalityTable,InterestRate,1)",15.621204152331885
Calculation,E17,"=B17-P_xt*D17+gamma2*(C17-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D17)",0.017085684348174136
Calculation,F17,=SumInsured*E17,"""1708.5684"""
Calculation,G17,=B17+gamma3*C17,0.6938026028162774
Calculation,H17,"=F17+alpha*t*GrossAnnualPrem*Act_axn_k(x+A17,MAX(5-A17,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""3415.4079"""
Calculation,I17,"=IF(AND(x+A17>=MinAgeFlex,A17>=n-MinTermFlex),1,0)",0.0
Calculation,J17,"=IF(OR(A17>n,I17),0,MIN(150,MAX(50,1%*(SumInsured-F17))))","""150"""
Calculation,K17,"=MAX(0,H17-J17)","""3265.4079"""
Calculation,L17,"=IFERROR(IF(A17>n,0,IF(A17<t,H17/G17,SumInsured)),0)","""4922.7372"""
Calculation,A18,2,2.0
Calculation,B18,"=IF(A18<=n,act_nGrAx(x+$A18,MAX(0,n-$A18),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A18,Sex,MortalityTable,InterestRate),0)",0.6519724347833002
Calculation,C18,"=Act_axn_k(x+$A18,MAX(0,n-$A18),Sex,MortalityTable,InterestRate,1)",20.235317006171048
Calculation,D18,"=Act_axn_k(x+$A18,MAX(0,t-$A18),Sex,MortalityTable,InterestRate,1)",14.919192104308186
Calculation,E18,"=B18-P_xt*D18+gamma2*(C18-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D18)",0.05582986298950499
Calculation,F18,=SumInsured*E18,"""5582.9863"""
Calculation,G18,=B18+gamma3*C18,0.7025607272987279
Calculation,H18,"=F18+alpha*t*GrossAnnualPrem*Act_axn_k(x+A18,MAX(5-A18,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""6875.8217"""
Calculation,I18,"=IF(AND(x+A18>=MinAgeFlex,A18>=n-MinTermFlex),1,0)",0.0
Calculation,J18,"=IF(OR(A18>n,I18),0,MIN(150,MAX(50,1%*(SumInsured-F18))))","""150"""
Calculation,K18,"=MAX(0,H18-J18)","""6725.8217"""
Calculation,L18,"=IFERROR(IF(A18>n,0,IF(A18<t,H18/G18,SumInsured)),0)","""9786.8006"""
Calculation,A19,3,3.0
Calculation,B19,"=IF(A19<=n,act_nGrAx(x+$A19,MAX(0,n-$A19),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A19,Sex,MortalityTable,InterestRate),0)",0.6623395947209115
Calculation,C19,"=Act_axn_k(x+$A19,MAX(0,n-$A19),Sex,MortalityTable,InterestRate,1)",19.63254070694141
Calculation,D19,"=Act_axn_k(x+$A19,MAX(0,t-$A19),Sex,MortalityTable,InterestRate,1)",14.20663384481256
Calculation,E19,"=B19-P_xt*D19+gamma2*(C19-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D19)",0.09512410516648707
Calculation,F19,=SumInsured*E19,"""9512.4105"""
Calculation,G19,=B19+gamma3*C19,0.711420946488265
Calculation,H19,"=F19+alpha*t*GrossAnnualPrem*Act_axn_k(x+A19,MAX(5-A19,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""10383.0348"""
Calculation,I19,"=IF(AND(x+A19>=MinAgeFlex,A19>=n-MinTermFlex),1,0)",0.0
Calculation,J19,"=IF(OR(A19>n,I19),0,MIN(150,MAX(50,1%*(SumInsured-F19))))","""150"""
Calculation,K19,"=MAX(0,H19-J19)","""10233.0348"""
Calculation,L19,"=IFERROR(IF(A19>n,0,IF(A19<t,H19/G19,SumInsured)),0)","""14594.7836"""
Calculation,A20,4,4.0
Calculation,B20,"=IF(A20<=n,act_nGrAx(x+$A20,MAX(0,n-$A20),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A20,Sex,MortalityTable,InterestRate),0)",0.6728223870534769
Calculation,C20,"=Act_axn_k(x+$A20,MAX(0,n-$A20),Sex,MortalityTable,InterestRate,1)",19.023041209890845
Calculation,D20,"=Act_axn_k(x+$A20,MAX(0,t-$A20),Sex,MortalityTable,InterestRate,1)",13.483418274793564
Calculation,E20,"=B20-P_xt*D20+gamma2*(C20-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D20)",0.13496949013199303
Calculation,F20,=SumInsured*E20,"""13496.949"""
Calculation,G20,=B20+gamma3*C20,0.720379990078204
Calculation,H20,"=F20+alpha*t*GrossAnnualPrem*Act_axn_k(x+A20,MAX(5-A20,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""13936.7755"""
Calculation,I20,"=IF(AND(x+A20>=MinAgeFlex,A20>=n-MinTermFlex),1,0)",0.0
Calculation,J20,"=IF(OR(A20>n,I20),0,MIN(150,MAX(50,1%*(SumInsured-F20))))","""150"""
Calculation,K20,"=MAX(0,H20-J20)","""13786.7755"""
Calculation,L20,"=IFERROR(IF(A20>n,0,IF(A20<t,H20/G20,SumInsured)),0)","""19346.4222"""
Calculation,A21,5,5.0
Calculation,B21,"=IF(A21<=n,act_nGrAx(x+$A21,MAX(0,n-$A21),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A21,Sex,MortalityTable,InterestRate),0)",0.683417191281626
Calculation,C21,"=Act_axn_k(x+$A21,MAX(0,n-$A21),Sex,MortalityTable,InterestRate,1)",18.40702902119689
Calculation,D21,"=Act_axn_k(x+$A21,MAX(0,t-$A21),Sex,MortalityTable,InterestRate,1)",12.749382293026253
Calculation,E21,"=B21-P_xt*D21+gamma2*(C21-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D21)",0.17536945425788478
Calculation,F21,=SumInsured*E21,"""17536.9454"""
Calculation,G21,=B21+gamma3*C21,0.7294347638346182
Calculation,H21,"=F21+alpha*t*GrossAnnualPrem*Act_axn_k(x+A21,MAX(5-A21,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""17536.9454"""
Calculation,I21,"=IF(AND(x+A21>=MinAgeFlex,A21>=n-MinTermFlex),1,0)",0.0
Calculation,J21,"=IF(OR(A21>n,I21),0,MIN(150,MAX(50,1%*(SumInsured-F21))))","""150"""
Calculation,K21,"=MAX(0,H21-J21)","""17386.9454"""
Calculation,L21,"=IFERROR(IF(A21>n,0,IF(A21<t,H21/G21,SumInsured)),0)","""24041.8284"""
Calculation,A22,6,6.0
Calculation,B22,"=IF(A22<=n,act_nGrAx(x+$A22,MAX(0,n-$A22),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A22,Sex,MortalityTable,InterestRate),0)",0.6941228960026647
Calculation,C22,"=Act_axn_k(x+$A22,MAX(0,n-$A22),Sex,MortalityTable,InterestRate,1)",17.7845687609879
Calculation,D22,"=Act_axn_k(x+$A22,MAX(0,t-$A22),Sex,MortalityTable,InterestRate,1)",12.00421375956845
Calculation,E22,"=B22-P_xt*D22+gamma2*(C22-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D22)",0.21633596864084273
Calculation,F22,=SumInsured*E22,"""21633.5969"""
Calculation,G22,=B22+gamma3*C22,0.7385843179051345
Calculation,H22,"=F22+alpha*t*GrossAnnualPrem*Act_axn_k(x+A22,MAX(5-A22,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""21633.5969"""
Calculation,I22,"=IF(AND(x+A22>=MinAgeFlex,A22>=n-MinTermFlex),1,0)",0.0
Calculation,J22,"=IF(OR(A22>n,I22),0,MIN(150,MAX(50,1%*(SumInsured-F22))))","""150"""
Calculation,K22,"=MAX(0,H22-J22)","""21483.5969"""
Calculation,L22,"=IFERROR(IF(A22>n,0,IF(A22<t,H22/G22,SumInsured)),0)","""29290.6258"""
Calculation,A23,7,7.0
Calculation,B23,"=IF(A23<=n,act_nGrAx(x+$A23,MAX(0,n-$A23),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A23,Sex,MortalityTable,InterestRate),0)",0.7049357662175477
Calculation,C23,"=Act_axn_k(x+$A23,MAX(0,n-$A23),Sex,MortalityTable,InterestRate,1)",17.15587759277976
Calculation,D23,"=Act_axn_k(x+$A23,MAX(0,t-$A23),Sex,MortalityTable,InterestRate,1)",11.247649370815745
Calculation,E23,"=B23-P_xt*D23+gamma2*(C23-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D23)",0.25787653725554543
Calculation,F23,=SumInsured*E23,"""25787.6537"""
Calculation,G23,=B23+gamma3*C23,0.7478254601994971
Calculation,H23,"=F23+alpha*t*GrossAnnualPrem*Act_axn_k(x+A23,MAX(5-A23,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""25787.6537"""
Calculation,I23,"=IF(AND(x+A23>=MinAgeFlex,A23>=n-MinTermFlex),1,0)",0.0
Calculation,J23,"=IF(OR(A23>n,I23),0,MIN(150,MAX(50,1%*(SumInsured-F23))))","""150"""
Calculation,K23,"=MAX(0,H23-J23)","""25637.6537"""
Calculation,L23,"=IFERROR(IF(A23>n,0,IF(A23<t,H23/G23,SumInsured)),0)","""34483.5193"""
Calculation,A24,8,8.0
Calculation,B24,"=IF(A24<=n,act_nGrAx(x+$A24,MAX(0,n-$A24),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A24,Sex,MortalityTable,InterestRate),0)",0.7158562538392343
Calculation,C24,"=Act_axn_k(x+$A24,MAX(0,n-$A24),Sex,MortalityTable,InterestRate,1)",16.52092924106161
Calculation,D24,"=Act_axn_k(x+$A24,MAX(0,t-$A24),Sex,MortalityTable,InterestRate,1)",10.479201093855696
Calculation,E24,"=B24-P_xt*D24+gamma2*(C24-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D24)",0.30001190745438006
Calculation,F24,=SumInsured*E24,"""30001.1907"""
Calculation,G24,=B24+gamma3*C24,0.7571585769418884
Calculation,H24,"=F24+alpha*t*GrossAnnualPrem*Act_axn_k(x+A24,MAX(5-A24,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""30001.1907"""
Calculation,I24,"=IF(AND(x+A24>=MinAgeFlex,A24>=n-MinTermFlex),1,0)",0.0
Calculation,J24,"=IF(OR(A24>n,I24),0,MIN(150,MAX(50,1%*(SumInsured-F24))))","""150"""
Calculation,K24,"=MAX(0,H24-J24)","""29851.1907"""
Calculation,L24,"=IFERROR(IF(A24>n,0,IF(A24<t,H24/G24,SumInsured)),0)","""39623.3915"""
Calculation,A25,9,9.0
Calculation,B25,"=IF(A25<=n,act_nGrAx(x+$A25,MAX(0,n-$A25),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A25,Sex,MortalityTable,InterestRate),0)",0.7268793899613282
Calculation,C25,"=Act_axn_k(x+$A25,MAX(0,n-$A25),Sex,MortalityTable,InterestRate,1)",15.880012612248448
Calculation,D25,"=Act_axn_k(x+$A25,MAX(0,t-$A25),Sex,MortalityTable,InterestRate,1)",9.69850648672706
Calculation,E25,"=B25-P_xt*D25+gamma2*(C25-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D25)",0.34275256844383917
Calculation,F25,=SumInsured*E25,"""34275.2568"""
Calculation,G25,=B25+gamma3*C25,0.7665794214919492
Calculation,H25,"=F25+alpha*t*GrossAnnualPrem*Act_axn_k(x+A25,MAX(5-A25,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""34275.2568"""
Calculation,I25,"=IF(AND(x+A25>=MinAgeFlex,A25>=n-MinTermFlex),1,0)",0.0
Calculation,J25,"=IF(OR(A25>n,I25),0,MIN(150,MAX(50,1%*(SumInsured-F25))))","""150"""
Calculation,K25,"=MAX(0,H25-J25)","""34125.2568"""
Calculation,L25,"=IFERROR(IF(A25>n,0,IF(A25<t,H25/G25,SumInsured)),0)","""44711.9449"""
Calculation,A26,10,10.0
Calculation,B26,"=IF(A26<=n,act_nGrAx(x+$A26,MAX(0,n-$A26),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A26,Sex,MortalityTable,InterestRate),0)",0.7380031706073322
Calculation,C26,"=Act_axn_k(x+$A26,MAX(0,n-$A26),Sex,MortalityTable,InterestRate,1)",15.233244223259339
Calculation,D26,"=Act_axn_k(x+$A26,MAX(0,t-$A26),Sex,MortalityTable,InterestRate,1)",8.904997404427364
Calculation,E26,"=B26-P_xt*D26+gamma2*(C26-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D26)",0.38612032712713734
Calculation,F26,=SumInsured*E26,"""38612.0327"""
Calculation,G26,=B26+gamma3*C26,0.7760862811654806
Calculation,H26,"=F26+alpha*t*GrossAnnualPrem*Act_axn_k(x+A26,MAX(5-A26,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""38612.0327"""
Calculation,I26,"=IF(AND(x+A26>=MinAgeFlex,A26>=n-MinTermFlex),1,0)",0.0
Calculation,J26,"=IF(OR(A26>n,I26),0,MIN(150,MAX(50,1%*(SumInsured-F26))))","""150"""
Calculation,K26,"=MAX(0,H26-J26)","""38462.0327"""
Calculation,L26,"=IFERROR(IF(A26>n,0,IF(A26<t,H26/G26,SumInsured)),0)","""49752.2423"""
Calculation,A27,11,11.0
Calculation,B27,"=IF(A27<=n,act_nGrAx(x+$A27,MAX(0,n-$A27),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A27,Sex,MortalityTable,InterestRate),0)",0.7492252457268609
Calculation,C27,"=Act_axn_k(x+$A27,MAX(0,n-$A27),Sex,MortalityTable,InterestRate,1)",14.58076071273807
Calculation,D27,"=Act_axn_k(x+$A27,MAX(0,t-$A27),Sex,MortalityTable,InterestRate,1)",8.098004487298601
Calculation,E27,"=B27-P_xt*D27+gamma2*(C27-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D27)",0.4301408854314264
Calculation,F27,=SumInsured*E27,"""43014.0885"""
Calculation,G27,=B27+gamma3*C27,0.7856771475087061
Calculation,H27,"=F27+alpha*t*GrossAnnualPrem*Act_axn_k(x+A27,MAX(5-A27,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""43014.0885"""
Calculation,I27,"=IF(AND(x+A27>=MinAgeFlex,A27>=n-MinTermFlex),1,0)",0.0
Calculation,J27,"=IF(OR(A27>n,I27),0,MIN(150,MAX(50,1%*(SumInsured-F27))))","""150"""
Calculation,K27,"=MAX(0,H27-J27)","""42864.0885"""
Calculation,L27,"=IFERROR(IF(A27>n,0,IF(A27<t,H27/G27,SumInsured)),0)","""54747.7914"""
Calculation,A28,12,12.0
Calculation,B28,"=IF(A28<=n,act_nGrAx(x+$A28,MAX(0,n-$A28),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A28,Sex,MortalityTable,InterestRate),0)",0.7605443620772302
Calculation,C28,"=Act_axn_k(x+$A28,MAX(0,n-$A28),Sex,MortalityTable,InterestRate,1)",13.922634947795249
Calculation,D28,"=Act_axn_k(x+$A28,MAX(0,t-$A28),Sex,MortalityTable,InterestRate,1)",7.276685557222148
Calculation,E28,"=B28-P_xt*D28+gamma2*(C28-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D28)",0.4748481608726271
Calculation,F28,=SumInsured*E28,"""47484.8161"""
Calculation,G28,=B28+gamma3*C28,0.7953509494467184
Calculation,H28,"=F28+alpha*t*GrossAnnualPrem*Act_axn_k(x+A28,MAX(5-A28,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""47484.8161"""
Calculation,I28,"=IF(AND(x+A28>=MinAgeFlex,A28>=n-MinTermFlex),1,0)",0.0
Calculation,J28,"=IF(OR(A28>n,I28),0,MIN(150,MAX(50,1%*(SumInsured-F28))))","""150"""
Calculation,K28,"=MAX(0,H28-J28)","""47334.8161"""
Calculation,L28,"=IFERROR(IF(A28>n,0,IF(A28<t,H28/G28,SumInsured)),0)","""59702.9728"""
Calculation,A29,13,13.0
Calculation,B29,"=IF(A29<=n,act_nGrAx(x+$A29,MAX(0,n-$A29),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A29,Sex,MortalityTable,InterestRate),0)",0.7719607061964255
Calculation,C29,"=Act_axn_k(x+$A29,MAX(0,n-$A29),Sex,MortalityTable,InterestRate,1)",13.258856082579262
Calculation,D29,"=Act_axn_k(x+$A29,MAX(0,t-$A29),Sex,MortalityTable,InterestRate,1)",6.439992371138744
Calculation,E29,"=B29-P_xt*D29+gamma2*(C29-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D29)",0.5202859894341998
Calculation,F29,=SumInsured*E29,"""52028.5989"""
Calculation,G29,=B29+gamma3*C29,0.8051078464028737
Calculation,H29,"=F29+alpha*t*GrossAnnualPrem*Act_axn_k(x+A29,MAX(5-A29,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""52028.5989"""
Calculation,I29,"=IF(AND(x+A29>=MinAgeFlex,A29>=n-MinTermFlex),1,0)",0.0
Calculation,J29,"=IF(OR(A29>n,I29),0,MIN(150,MAX(50,1%*(SumInsured-F29))))","""150"""
Calculation,K29,"=MAX(0,H29-J29)","""51878.5989"""
Calculation,L29,"=IFERROR(IF(A29>n,0,IF(A29<t,H29/G29,SumInsured)),0)","""64623.1423"""
Calculation,A30,14,14.0
Calculation,B30,"=IF(A30<=n,act_nGrAx(x+$A30,MAX(0,n-$A30),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A30,Sex,MortalityTable,InterestRate),0)",0.7834747382680015
Calculation,C30,"=Act_axn_k(x+$A30,MAX(0,n-$A30),Sex,MortalityTable,InterestRate,1)",12.589397360703279
Calculation,D30,"=Act_axn_k(x+$A30,MAX(0,t-$A30),Sex,MortalityTable,InterestRate,1)",5.58667343332173
Calculation,E30,"=B30-P_xt*D30+gamma2*(C30-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D30)",0.5665069270440265
Calculation,F30,=SumInsured*E30,"""56650.6927"""
Calculation,G30,=B30+gamma3*C30,0.8149482316697597
Calculation,H30,"=F30+alpha*t*GrossAnnualPrem*Act_axn_k(x+A30,MAX(5-A30,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""56650.6927"""
Calculation,I30,"=IF(AND(x+A30>=MinAgeFlex,A30>=n-MinTermFlex),1,0)",0.0
Calculation,J30,"=IF(OR(A30>n,I30),0,MIN(150,MAX(50,1%*(SumInsured-F30))))","""150"""
Calculation,K30,"=MAX(0,H30-J30)","""56500.6927"""
Calculation,L30,"=IFERROR(IF(A30>n,0,IF(A30<t,H30/G30,SumInsured)),0)","""69514.4679"""
Calculation,A31,15,15.0
Calculation,B31,"=IF(A31<=n,act_nGrAx(x+$A31,MAX(0,n-$A31),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A31,Sex,MortalityTable,InterestRate),0)",0.795096555571746
Calculation,C31,"=Act_axn_k(x+$A31,MAX(0,n-$A31),Sex,MortalityTable,InterestRate,1)",11.913671697471306
Calculation,D31,"=Act_axn_k(x+$A31,MAX(0,t-$A31),Sex,MortalityTable,InterestRate,1)",4.715009742732018
Calculation,E31,"=B31-P_xt*D31+gamma2*(C31-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D31)",0.613591939553059
Calculation,F31,=SumInsured*E31,"""61359.194"""
Calculation,G31,=B31+gamma3*C31,0.8248807348154242
Calculation,H31,"=F31+alpha*t*GrossAnnualPrem*Act_axn_k(x+A31,MAX(5-A31,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""61359.194"""
Calculation,I31,"=IF(AND(x+A31>=MinAgeFlex,A31>=n-MinTermFlex),1,0)",0.0
Calculation,J31,"=IF(OR(A31>n,I31),0,MIN(150,MAX(50,1%*(SumInsured-F31))))","""150"""
Calculation,K31,"=MAX(0,H31-J31)","""61209.194"""
Calculation,L31,"=IFERROR(IF(A31>n,0,IF(A31<t,H31/G31,SumInsured)),0)","""74385.534"""
Calculation,A32,16,16.0
Calculation,B32,"=IF(A32<=n,act_nGrAx(x+$A32,MAX(0,n-$A32),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A32,Sex,MortalityTable,InterestRate),0)",0.8068404040744293
Calculation,C32,"=Act_axn_k(x+$A32,MAX(0,n-$A32),Sex,MortalityTable,InterestRate,1)",11.230850791672268
Calculation,D32,"=Act_axn_k(x+$A32,MAX(0,t-$A32),Sex,MortalityTable,InterestRate,1)",3.8229773871518575
Calculation,E32,"=B32-P_xt*D32+gamma2*(C32-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D32)",0.6616385405994291
Calculation,F32,=SumInsured*E32,"""66163.8541"""
Calculation,G32,=B32+gamma3*C32,0.8349175310536099
Calculation,H32,"=F32+alpha*t*GrossAnnualPrem*Act_axn_k(x+A32,MAX(5-A32,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""66163.8541"""
Calculation,I32,"=IF(AND(x+A32>=MinAgeFlex,A32>=n-MinTermFlex),1,0)",0.0
Calculation,J32,"=IF(OR(A32>n,I32),0,MIN(150,MAX(50,1%*(SumInsured-F32))))","""150"""
Calculation,K32,"=MAX(0,H32-J32)","""66013.8541"""
Calculation,L32,"=IFERROR(IF(A32>n,0,IF(A32<t,H32/G32,SumInsured)),0)","""79245.9753"""
Calculation,A33,17,17.0
Calculation,B33,"=IF(A33<=n,act_nGrAx(x+$A33,MAX(0,n-$A33),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A33,Sex,MortalityTable,InterestRate),0)",0.8187231548769123
Calculation,C33,"=Act_axn_k(x+$A33,MAX(0,n-$A33),Sex,MortalityTable,InterestRate,1)",10.539953709299303
Calculation,D33,"=Act_axn_k(x+$A33,MAX(0,t-$A33),Sex,MortalityTable,InterestRate,1)",2.90826751253167
Calculation,E33,"=B33-P_xt*D33+gamma2*(C33-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D33)",0.7107585469918153
Calculation,F33,=SumInsured*E33,"""71075.8547"""
Calculation,G33,=B33+gamma3*C33,0.8450730391501605
Calculation,H33,"=F33+alpha*t*GrossAnnualPrem*Act_axn_k(x+A33,MAX(5-A33,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""71075.8547"""
Calculation,I33,"=IF(AND(x+A33>=MinAgeFlex,A33>=n-MinTermFlex),1,0)",0.0
Calculation,J33,"=IF(OR(A33>n,I33),0,MIN(150,MAX(50,1%*(SumInsured-F33))))","""150"""
Calculation,K33,"=MAX(0,H33-J33)","""70925.8547"""
Calculation,L33,"=IFERROR(IF(A33>n,0,IF(A33<t,H33/G33,SumInsured)),0)","""84106.1676"""
Calculation,A34,18,18.0
Calculation,B34,"=IF(A34<=n,act_nGrAx(x+$A34,MAX(0,n-$A34),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A34,Sex,MortalityTable,InterestRate),0)",0.830762893646463
Calculation,C34,"=Act_axn_k(x+$A34,MAX(0,n-$A34),Sex,MortalityTable,InterestRate,1)",9.839928897983867
Calculation,D34,"=Act_axn_k(x+$A34,MAX(0,t-$A34),Sex,MortalityTable,InterestRate,1)",1.9682712530712323
Calculation,E34,"=B34-P_xt*D34+gamma2*(C34-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D34)",0.7610773983748185
Calculation,F34,=SumInsured*E34,"""76107.7398"""
Calculation,G34,=B34+gamma3*C34,0.8553627158914228
Calculation,H34,"=F34+alpha*t*GrossAnnualPrem*Act_axn_k(x+A34,MAX(5-A34,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""76107.7398"""
Calculation,I34,"=IF(AND(x+A34>=MinAgeFlex,A34>=n-MinTermFlex),1,0)",0.0
Calculation,J34,"=IF(OR(A34>n,I34),0,MIN(150,MAX(50,1%*(SumInsured-F34))))","""150"""
Calculation,K34,"=MAX(0,H34-J34)","""75957.7398"""
Calculation,L34,"=IFERROR(IF(A34>n,0,IF(A34<t,H34/G34,SumInsured)),0)","""88977.1537"""
Calculation,A35,19,19.0
Calculation,B35,"=IF(A35<=n,act_nGrAx(x+$A35,MAX(0,n-$A35),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A35,Sex,MortalityTable,InterestRate),0)",0.8429798585135433
Calculation,C35,"=Act_axn_k(x+$A35,MAX(0,n-$A35),Sex,MortalityTable,InterestRate,1)",9.129599654998112
Calculation,D35,"=Act_axn_k(x+$A35,MAX(0,t-$A35),Sex,MortalityTable,InterestRate,1)",0.9999999999999972
Calculation,E35,"=B35-P_xt*D35+gamma2*(C35-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D35)",0.8127383480459093
Calculation,F35,=SumInsured*E35,"""81273.8348"""
Calculation,G35,=B35+gamma3*C35,0.8658038576510386
Calculation,H35,"=F35+alpha*t*GrossAnnualPrem*Act_axn_k(x+A35,MAX(5-A35,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""81273.8348"""
Calculation,I35,"=IF(AND(x+A35>=MinAgeFlex,A35>=n-MinTermFlex),1,0)",0.0
Calculation,J35,"=IF(OR(A35>n,I35),0,MIN(150,MAX(50,1%*(SumInsured-F35))))","""150"""
Calculation,K35,"=MAX(0,H35-J35)","""81123.8348"""
Calculation,L35,"=IFERROR(IF(A35>n,0,IF(A35<t,H35/G35,SumInsured)),0)","""93870.9548"""
Calculation,A36,20,20.0
Calculation,B36,"=IF(A36<=n,act_nGrAx(x+$A36,MAX(0,n-$A36),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A36,Sex,MortalityTable,InterestRate),0)",0.8553966621309441
Calculation,C36,"=Act_axn_k(x+$A36,MAX(0,n-$A36),Sex,MortalityTable,InterestRate,1)",8.407651216100595
Calculation,D36,"=Act_axn_k(x+$A36,MAX(0,t-$A36),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E36,"=B36-P_xt*D36+gamma2*(C36-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D36)",0.8659062261510698
Calculation,F36,=SumInsured*E36,"""86590.6226"""
Calculation,G36,=B36+gamma3*C36,0.8764157901711956
Calculation,H36,"=F36+alpha*t*GrossAnnualPrem*Act_axn_k(x+A36,MAX(5-A36,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""86590.6226"""
Calculation,I36,"=IF(AND(x+A36>=MinAgeFlex,A36>=n-MinTermFlex),1,0)",0.0
Calculation,J36,"=IF(OR(A36>n,I36),0,MIN(150,MAX(50,1%*(SumInsured-F36))))","""134.0938"""
Calculation,K36,"=MAX(0,H36-J36)","""86456.5288"""
Calculation,L36,"=IFERROR(IF(A36>n,0,IF(A36<t,H36/G36,SumInsured)),0)","""100000"""
Calculation,A37,21,21.0
Calculation,B37,"=IF(A37<=n,act_nGrAx(x+$A37,MAX(0,n-$A37),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A37,Sex,MortalityTable,InterestRate),0)",0.8680403142570174
Calculation,C37,"=Act_axn_k(x+$A37,MAX(0,n-$A37),Sex,MortalityTable,InterestRate,1)",7.672513156770438
Calculation,D37,"=Act_axn_k(x+$A37,MAX(0,t-$A37),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E37,"=B37-P_xt*D37+gamma2*(C37-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D37)",0.8776309557029804
Calculation,F37,=SumInsured*E37,"""87763.0956"""
Calculation,G37,=B37+gamma3*C37,0.8872215971489436
Calculation,H37,"=F37+alpha*t*GrossAnnualPrem*Act_axn_k(x+A37,MAX(5-A37,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""87763.0956"""
Calculation,I37,"=IF(AND(x+A37>=MinAgeFlex,A37>=n-MinTermFlex),1,0)",0.0
Calculation,J37,"=IF(OR(A37>n,I37),0,MIN(150,MAX(50,1%*(SumInsured-F37))))","""122.369"""
Calculation,K37,"=MAX(0,H37-J37)","""87640.7265"""
Calculation,L37,"=IFERROR(IF(A37>n,0,IF(A37<t,H37/G37,SumInsured)),0)","""100000"""
Calculation,A38,22,22.0
Calculation,B38,"=IF(A38<=n,act_nGrAx(x+$A38,MAX(0,n-$A38),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A38,Sex,MortalityTable,InterestRate),0)",0.8809423750317524
Calculation,C38,"=Act_axn_k(x+$A38,MAX(0,n-$A38),Sex,MortalityTable,InterestRate,1)",6.9223504802966715
Calculation,D38,"=Act_axn_k(x+$A38,MAX(0,t-$A38),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E38,"=B38-P_xt*D38+gamma2*(C38-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D38)",0.8895953131321233
Calculation,F38,=SumInsured*E38,"""88959.5313"""
Calculation,G38,=B38+gamma3*C38,0.8982482512324941
Calculation,H38,"=F38+alpha*t*GrossAnnualPrem*Act_axn_k(x+A38,MAX(5-A38,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""88959.5313"""
Calculation,I38,"=IF(AND(x+A38>=MinAgeFlex,A38>=n-MinTermFlex),1,0)",0.0
Calculation,J38,"=IF(OR(A38>n,I38),0,MIN(150,MAX(50,1%*(SumInsured-F38))))","""110.4047"""
Calculation,K38,"=MAX(0,H38-J38)","""88849.1266"""
Calculation,L38,"=IFERROR(IF(A38>n,0,IF(A38<t,H38/G38,SumInsured)),0)","""100000"""
Calculation,A39,23,23.0
Calculation,B39,"=IF(A39<=n,act_nGrAx(x+$A39,MAX(0,n-$A39),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A39,Sex,MortalityTable,InterestRate),0)",0.894140474375828
Calculation,C39,"=Act_axn_k(x+$A39,MAX(0,n-$A39),Sex,MortalityTable,InterestRate,1)",6.154975275576831
Calculation,D39,"=Act_axn_k(x+$A39,MAX(0,t-$A39),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E39,"=B39-P_xt*D39+gamma2*(C39-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D39)",0.901834193470299
Calculation,F39,=SumInsured*E39,"""90183.4193"""
Calculation,G39,=B39+gamma3*C39,0.9095279125647701
Calculation,H39,"=F39+alpha*t*GrossAnnualPrem*Act_axn_k(x+A39,MAX(5-A39,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""90183.4193"""
Calculation,I39,"=IF(AND(x+A39>=MinAgeFlex,A39>=n-MinTermFlex),1,0)",0.0
Calculation,J39,"=IF(OR(A39>n,I39),0,MIN(150,MAX(50,1%*(SumInsured-F39))))","""98.1658"""
Calculation,K39,"=MAX(0,H39-J39)","""90085.2535"""
Calculation,L39,"=IFERROR(IF(A39>n,0,IF(A39<t,H39/G39,SumInsured)),0)","""100000"""
Calculation,A40,24,24.0
Calculation,B40,"=IF(A40<=n,act_nGrAx(x+$A40,MAX(0,n-$A40),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A40,Sex,MortalityTable,InterestRate),0)",0.9076799898864822
Calculation,C40,"=Act_axn_k(x+$A40,MAX(0,n-$A40),Sex,MortalityTable,InterestRate,1)",5.367749159457335
Calculation,D40,"=Act_axn_k(x+$A40,MAX(0,t-$A40),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E40,"=B40-P_xt*D40+gamma2*(C40-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D40)",0.9143896763358039
Calculation,F40,=SumInsured*E40,"""91438.9676"""
Calculation,G40,=B40+gamma3*C40,0.9210993627851256
Calculation,H40,"=F40+alpha*t*GrossAnnualPrem*Act_axn_k(x+A40,MAX(5-A40,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""91438.9676"""
Calculation,I40,"=IF(AND(x+A40>=MinAgeFlex,A40>=n-MinTermFlex),1,0)",0.0
Calculation,J40,"=IF(OR(A40>n,I40),0,MIN(150,MAX(50,1%*(SumInsured-F40))))","""85.6103"""
Calculation,K40,"=MAX(0,H40-J40)","""91353.3573"""
Calculation,L40,"=IFERROR(IF(A40>n,0,IF(A40<t,H40/G40,SumInsured)),0)","""100000"""
Calculation,A41,25,25.0
Calculation,B41,"=IF(A41<=n,act_nGrAx(x+$A41,MAX(0,n-$A41),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A41,Sex,MortalityTable,InterestRate),0)",0.9216159182042138
Calculation,C41,"=Act_axn_k(x+$A41,MAX(0,n-$A41),Sex,MortalityTable,InterestRate,1)",4.557474470126218
Calculation,D41,"=Act_axn_k(x+$A41,MAX(0,t-$A41),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E41,"=B41-P_xt*D41+gamma2*(C41-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D41)",0.9273127612918716
Calculation,F41,=SumInsured*E41,"""92731.2761"""
Calculation,G41,=B41+gamma3*C41,0.9330096043795294
Calculation,H41,"=F41+alpha*t*GrossAnnualPrem*Act_axn_k(x+A41,MAX(5-A41,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""92731.2761"""
Calculation,I41,"=IF(AND(x+A41>=MinAgeFlex,A41>=n-MinTermFlex),1,0)",1.0
Calculation,J41,"=IF(OR(A41>n,I41),0,MIN(150,MAX(50,1%*(SumInsured-F41))))","""0"""
Calculation,K41,"=MAX(0,H41-J41)","""92731.2761"""
Calculation,L41,"=IFERROR(IF(A41>n,0,IF(A41<t,H41/G41,SumInsured)),0)","""100000"""
Calculation,A42,26,26.0
Calculation,B42,"=IF(A42<=n,act_nGrAx(x+$A42,MAX(0,n-$A42),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A42,Sex,MortalityTable,InterestRate),0)",0.936011845465061
Calculation,C42,"=Act_axn_k(x+$A42,MAX(0,n-$A42),Sex,MortalityTable,InterestRate,1)",3.7204541279596572
Calculation,D42,"=Act_axn_k(x+$A42,MAX(0,t-$A42),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E42,"=B42-P_xt*D42+gamma2*(C42-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D42)",0.9406624131250105
Calculation,F42,=SumInsured*E42,"""94066.2413"""
Calculation,G42,=B42+gamma3*C42,0.9453129807849602
Calculation,H42,"=F42+alpha*t*GrossAnnualPrem*Act_axn_k(x+A42,MAX(5-A42,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""94066.2413"""
Calculation,I42,"=IF(AND(x+A42>=MinAgeFlex,A42>=n-MinTermFlex),1,0)",1.0
Calculation,J42,"=IF(OR(A42>n,I42),0,MIN(150,MAX(50,1%*(SumInsured-F42))))","""0"""
Calculation,K42,"=MAX(0,H42-J42)","""94066.2413"""
Calculation,L42,"=IFERROR(IF(A42>n,0,IF(A42<t,H42/G42,SumInsured)),0)","""100000"""
Calculation,A43,27,27.0
Calculation,B43,"=IF(A43<=n,act_nGrAx(x+$A43,MAX(0,n-$A43),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A43,Sex,MortalityTable,InterestRate),0)",0.9509422995984339
Calculation,C43,"=Act_axn_k(x+$A43,MAX(0,n-$A43),Sex,MortalityTable,InterestRate,1)",2.8523548662050255
Calculation,D43,"=Act_axn_k(x+$A43,MAX(0,t-$A43),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E43,"=B43-P_xt*D43+gamma2*(C43-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D43)",0.9545077431811901
Calculation,F43,=SumInsured*E43,"""95450.7743"""
Calculation,G43,=B43+gamma3*C43,0.9580731867639465
Calculation,H43,"=F43+alpha*t*GrossAnnualPrem*Act_axn_k(x+A43,MAX(5-A43,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""95450.7743"""
Calculation,I43,"=IF(AND(x+A43>=MinAgeFlex,A43>=n-MinTermFlex),1,0)",1.0
Calculation,J43,"=IF(OR(A43>n,I43),0,MIN(150,MAX(50,1%*(SumInsured-F43))))","""0"""
Calculation,K43,"=MAX(0,H43-J43)","""95450.7743"""
Calculation,L43,"=IFERROR(IF(A43>n,0,IF(A43<t,H43/G43,SumInsured)),0)","""100000"""
Calculation,A44,28,28.0
Calculation,B44,"=IF(A44<=n,act_nGrAx(x+$A44,MAX(0,n-$A44),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A44,Sex,MortalityTable,InterestRate),0)",0.9665000670091556
Calculation,C44,"=Act_axn_k(x+$A44,MAX(0,n-$A44),Sex,MortalityTable,InterestRate,1)",1.9477818181817932
Calculation,D44,"=Act_axn_k(x+$A44,MAX(0,t-$A44),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E44,"=B44-P_xt*D44+gamma2*(C44-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D44)",0.9689347942818828
Calculation,F44,=SumInsured*E44,"""96893.4794"""
Calculation,G44,=B44+gamma3*C44,0.9713695215546101
Calculation,H44,"=F44+alpha*t*GrossAnnualPrem*Act_axn_k(x+A44,MAX(5-A44,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""96893.4794"""
Calculation,I44,"=IF(AND(x+A44>=MinAgeFlex,A44>=n-MinTermFlex),1,0)",1.0
Calculation,J44,"=IF(OR(A44>n,I44),0,MIN(150,MAX(50,1%*(SumInsured-F44))))","""0"""
Calculation,K44,"=MAX(0,H44-J44)","""96893.4794"""
Calculation,L44,"=IFERROR(IF(A44>n,0,IF(A44<t,H44/G44,SumInsured)),0)","""100000"""
Calculation,A45,29,29.0
Calculation,B45,"=IF(A45<=n,act_nGrAx(x+$A45,MAX(0,n-$A45),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A45,Sex,MortalityTable,InterestRate),0)",0.9828009828009829
Calculation,C45,"=Act_axn_k(x+$A45,MAX(0,n-$A45),Sex,MortalityTable,InterestRate,1)",0.9999999999999992
Calculation,D45,"=Act_axn_k(x+$A45,MAX(0,t-$A45),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E45,"=B45-P_xt*D45+gamma2*(C45-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D45)",0.9840509828009829
Calculation,F45,=SumInsured*E45,"""98405.0983"""
Calculation,G45,=B45+gamma3*C45,0.9853009828009829
Calculation,H45,"=F45+alpha*t*GrossAnnualPrem*Act_axn_k(x+A45,MAX(5-A45,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""98405.0983"""
Calculation,I45,"=IF(AND(x+A45>=MinAgeFlex,A45>=n-MinTermFlex),1,0)",1.0
Calculation,J45,"=IF(OR(A45>n,I45),0,MIN(150,MAX(50,1%*(SumInsured-F45))))","""0"""
Calculation,K45,"=MAX(0,H45-J45)","""98405.0983"""
Calculation,L45,"=IFERROR(IF(A45>n,0,IF(A45<t,H45/G45,SumInsured)),0)","""100000"""
Calculation,A46,30,30.0
Calculation,B46,"=IF(A46<=n,act_nGrAx(x+$A46,MAX(0,n-$A46),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A46,Sex,MortalityTable,InterestRate),0)",1.0
Calculation,C46,"=Act_axn_k(x+$A46,MAX(0,n-$A46),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D46,"=Act_axn_k(x+$A46,MAX(0,t-$A46),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E46,"=B46-P_xt*D46+gamma2*(C46-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D46)",1.0
Calculation,F46,=SumInsured*E46,"""100000"""
Calculation,G46,=B46+gamma3*C46,1.0
Calculation,H46,"=F46+alpha*t*GrossAnnualPrem*Act_axn_k(x+A46,MAX(5-A46,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""100000"""
Calculation,I46,"=IF(AND(x+A46>=MinAgeFlex,A46>=n-MinTermFlex),1,0)",1.0
Calculation,J46,"=IF(OR(A46>n,I46),0,MIN(150,MAX(50,1%*(SumInsured-F46))))","""0"""
Calculation,K46,"=MAX(0,H46-J46)","""100000"""
Calculation,L46,"=IFERROR(IF(A46>n,0,IF(A46<t,H46/G46,SumInsured)),0)","""100000"""
Calculation,A47,31,31.0
Calculation,B47,"=IF(A47<=n,act_nGrAx(x+$A47,MAX(0,n-$A47),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A47,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C47,"=Act_axn_k(x+$A47,MAX(0,n-$A47),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D47,"=Act_axn_k(x+$A47,MAX(0,t-$A47),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E47,"=B47-P_xt*D47+gamma2*(C47-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D47)",0.0
Calculation,F47,=SumInsured*E47,"""0"""
Calculation,G47,=B47+gamma3*C47,0.0
Calculation,H47,"=F47+alpha*t*GrossAnnualPrem*Act_axn_k(x+A47,MAX(5-A47,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I47,"=IF(AND(x+A47>=MinAgeFlex,A47>=n-MinTermFlex),1,0)",1.0
Calculation,J47,"=IF(OR(A47>n,I47),0,MIN(150,MAX(50,1%*(SumInsured-F47))))","""0"""
Calculation,K47,"=MAX(0,H47-J47)","""0"""
Calculation,L47,"=IFERROR(IF(A47>n,0,IF(A47<t,H47/G47,SumInsured)),0)","""0"""
Calculation,A48,32,32.0
Calculation,B48,"=IF(A48<=n,act_nGrAx(x+$A48,MAX(0,n-$A48),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A48,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C48,"=Act_axn_k(x+$A48,MAX(0,n-$A48),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D48,"=Act_axn_k(x+$A48,MAX(0,t-$A48),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E48,"=B48-P_xt*D48+gamma2*(C48-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D48)",0.0
Calculation,F48,=SumInsured*E48,"""0"""
Calculation,G48,=B48+gamma3*C48,0.0
Calculation,H48,"=F48+alpha*t*GrossAnnualPrem*Act_axn_k(x+A48,MAX(5-A48,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I48,"=IF(AND(x+A48>=MinAgeFlex,A48>=n-MinTermFlex),1,0)",1.0
Calculation,J48,"=IF(OR(A48>n,I48),0,MIN(150,MAX(50,1%*(SumInsured-F48))))","""0"""
Calculation,K48,"=MAX(0,H48-J48)","""0"""
Calculation,L48,"=IFERROR(IF(A48>n,0,IF(A48<t,H48/G48,SumInsured)),0)","""0"""
Calculation,A49,33,33.0
Calculation,B49,"=IF(A49<=n,act_nGrAx(x+$A49,MAX(0,n-$A49),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A49,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C49,"=Act_axn_k(x+$A49,MAX(0,n-$A49),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D49,"=Act_axn_k(x+$A49,MAX(0,t-$A49),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E49,"=B49-P_xt*D49+gamma2*(C49-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D49)",0.0
Calculation,F49,=SumInsured*E49,"""0"""
Calculation,G49,=B49+gamma3*C49,0.0
Calculation,H49,"=F49+alpha*t*GrossAnnualPrem*Act_axn_k(x+A49,MAX(5-A49,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I49,"=IF(AND(x+A49>=MinAgeFlex,A49>=n-MinTermFlex),1,0)",1.0
Calculation,J49,"=IF(OR(A49>n,I49),0,MIN(150,MAX(50,1%*(SumInsured-F49))))","""0"""
Calculation,K49,"=MAX(0,H49-J49)","""0"""
Calculation,L49,"=IFERROR(IF(A49>n,0,IF(A49<t,H49/G49,SumInsured)),0)","""0"""
Calculation,A50,34,34.0
Calculation,B50,"=IF(A50<=n,act_nGrAx(x+$A50,MAX(0,n-$A50),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A50,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C50,"=Act_axn_k(x+$A50,MAX(0,n-$A50),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D50,"=Act_axn_k(x+$A50,MAX(0,t-$A50),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E50,"=B50-P_xt*D50+gamma2*(C50-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D50)",0.0
Calculation,F50,=SumInsured*E50,"""0"""
Calculation,G50,=B50+gamma3*C50,0.0
Calculation,H50,"=F50+alpha*t*GrossAnnualPrem*Act_axn_k(x+A50,MAX(5-A50,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I50,"=IF(AND(x+A50>=MinAgeFlex,A50>=n-MinTermFlex),1,0)",1.0
Calculation,J50,"=IF(OR(A50>n,I50),0,MIN(150,MAX(50,1%*(SumInsured-F50))))","""0"""
Calculation,K50,"=MAX(0,H50-J50)","""0"""
Calculation,L50,"=IFERROR(IF(A50>n,0,IF(A50<t,H50/G50,SumInsured)),0)","""0"""
Calculation,A51,35,35.0
Calculation,B51,"=IF(A51<=n,act_nGrAx(x+$A51,MAX(0,n-$A51),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A51,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C51,"=Act_axn_k(x+$A51,MAX(0,n-$A51),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D51,"=Act_axn_k(x+$A51,MAX(0,t-$A51),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E51,"=B51-P_xt*D51+gamma2*(C51-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D51)",0.0
Calculation,F51,=SumInsured*E51,"""0"""
Calculation,G51,=B51+gamma3*C51,0.0
Calculation,H51,"=F51+alpha*t*GrossAnnualPrem*Act_axn_k(x+A51,MAX(5-A51,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I51,"=IF(AND(x+A51>=MinAgeFlex,A51>=n-MinTermFlex),1,0)",1.0
Calculation,J51,"=IF(OR(A51>n,I51),0,MIN(150,MAX(50,1%*(SumInsured-F51))))","""0"""
Calculation,K51,"=MAX(0,H51-J51)","""0"""
Calculation,L51,"=IFERROR(IF(A51>n,0,IF(A51<t,H51/G51,SumInsured)),0)","""0"""
Calculation,A52,36,36.0
Calculation,B52,"=IF(A52<=n,act_nGrAx(x+$A52,MAX(0,n-$A52),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A52,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C52,"=Act_axn_k(x+$A52,MAX(0,n-$A52),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D52,"=Act_axn_k(x+$A52,MAX(0,t-$A52),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E52,"=B52-P_xt*D52+gamma2*(C52-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D52)",0.0
Calculation,F52,=SumInsured*E52,"""0"""
Calculation,G52,=B52+gamma3*C52,0.0
Calculation,H52,"=F52+alpha*t*GrossAnnualPrem*Act_axn_k(x+A52,MAX(5-A52,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I52,"=IF(AND(x+A52>=MinAgeFlex,A52>=n-MinTermFlex),1,0)",1.0
Calculation,J52,"=IF(OR(A52>n,I52),0,MIN(150,MAX(50,1%*(SumInsured-F52))))","""0"""
Calculation,K52,"=MAX(0,H52-J52)","""0"""
Calculation,L52,"=IFERROR(IF(A52>n,0,IF(A52<t,H52/G52,SumInsured)),0)","""0"""
Calculation,A53,37,37.0
Calculation,B53,"=IF(A53<=n,act_nGrAx(x+$A53,MAX(0,n-$A53),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A53,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C53,"=Act_axn_k(x+$A53,MAX(0,n-$A53),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D53,"=Act_axn_k(x+$A53,MAX(0,t-$A53),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E53,"=B53-P_xt*D53+gamma2*(C53-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D53)",0.0
Calculation,F53,=SumInsured*E53,"""0"""
Calculation,G53,=B53+gamma3*C53,0.0
Calculation,H53,"=F53+alpha*t*GrossAnnualPrem*Act_axn_k(x+A53,MAX(5-A53,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I53,"=IF(AND(x+A53>=MinAgeFlex,A53>=n-MinTermFlex),1,0)",1.0
Calculation,J53,"=IF(OR(A53>n,I53),0,MIN(150,MAX(50,1%*(SumInsured-F53))))","""0"""
Calculation,K53,"=MAX(0,H53-J53)","""0"""
Calculation,L53,"=IFERROR(IF(A53>n,0,IF(A53<t,H53/G53,SumInsured)),0)","""0"""
Calculation,A54,38,38.0
Calculation,B54,"=IF(A54<=n,act_nGrAx(x+$A54,MAX(0,n-$A54),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A54,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C54,"=Act_axn_k(x+$A54,MAX(0,n-$A54),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D54,"=Act_axn_k(x+$A54,MAX(0,t-$A54),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E54,"=B54-P_xt*D54+gamma2*(C54-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D54)",0.0
Calculation,F54,=SumInsured*E54,"""0"""
Calculation,G54,=B54+gamma3*C54,0.0
Calculation,H54,"=F54+alpha*t*GrossAnnualPrem*Act_axn_k(x+A54,MAX(5-A54,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I54,"=IF(AND(x+A54>=MinAgeFlex,A54>=n-MinTermFlex),1,0)",1.0
Calculation,J54,"=IF(OR(A54>n,I54),0,MIN(150,MAX(50,1%*(SumInsured-F54))))","""0"""
Calculation,K54,"=MAX(0,H54-J54)","""0"""
Calculation,L54,"=IFERROR(IF(A54>n,0,IF(A54<t,H54/G54,SumInsured)),0)","""0"""
Calculation,A55,39,39.0
Calculation,B55,"=IF(A55<=n,act_nGrAx(x+$A55,MAX(0,n-$A55),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A55,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C55,"=Act_axn_k(x+$A55,MAX(0,n-$A55),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D55,"=Act_axn_k(x+$A55,MAX(0,t-$A55),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E55,"=B55-P_xt*D55+gamma2*(C55-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D55)",0.0
Calculation,F55,=SumInsured*E55,"""0"""
Calculation,G55,=B55+gamma3*C55,0.0
Calculation,H55,"=F55+alpha*t*GrossAnnualPrem*Act_axn_k(x+A55,MAX(5-A55,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I55,"=IF(AND(x+A55>=MinAgeFlex,A55>=n-MinTermFlex),1,0)",1.0
Calculation,J55,"=IF(OR(A55>n,I55),0,MIN(150,MAX(50,1%*(SumInsured-F55))))","""0"""
Calculation,K55,"=MAX(0,H55-J55)","""0"""
Calculation,L55,"=IFERROR(IF(A55>n,0,IF(A55<t,H55/G55,SumInsured)),0)","""0"""
Calculation,A56,40,40.0
Calculation,B56,"=IF(A56<=n,act_nGrAx(x+$A56,MAX(0,n-$A56),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A56,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C56,"=Act_axn_k(x+$A56,MAX(0,n-$A56),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D56,"=Act_axn_k(x+$A56,MAX(0,t-$A56),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E56,"=B56-P_xt*D56+gamma2*(C56-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D56)",0.0
Calculation,F56,=SumInsured*E56,"""0"""
Calculation,G56,=B56+gamma3*C56,0.0
Calculation,H56,"=F56+alpha*t*GrossAnnualPrem*Act_axn_k(x+A56,MAX(5-A56,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I56,"=IF(AND(x+A56>=MinAgeFlex,A56>=n-MinTermFlex),1,0)",1.0
Calculation,J56,"=IF(OR(A56>n,I56),0,MIN(150,MAX(50,1%*(SumInsured-F56))))","""0"""
Calculation,K56,"=MAX(0,H56-J56)","""0"""
Calculation,L56,"=IFERROR(IF(A56>n,0,IF(A56<t,H56/G56,SumInsured)),0)","""0"""
Calculation,A57,41,41.0
Calculation,B57,"=IF(A57<=n,act_nGrAx(x+$A57,MAX(0,n-$A57),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A57,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C57,"=Act_axn_k(x+$A57,MAX(0,n-$A57),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D57,"=Act_axn_k(x+$A57,MAX(0,t-$A57),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E57,"=B57-P_xt*D57+gamma2*(C57-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D57)",0.0
Calculation,F57,=SumInsured*E57,"""0"""
Calculation,G57,=B57+gamma3*C57,0.0
Calculation,H57,"=F57+alpha*t*GrossAnnualPrem*Act_axn_k(x+A57,MAX(5-A57,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I57,"=IF(AND(x+A57>=MinAgeFlex,A57>=n-MinTermFlex),1,0)",1.0
Calculation,J57,"=IF(OR(A57>n,I57),0,MIN(150,MAX(50,1%*(SumInsured-F57))))","""0"""
Calculation,K57,"=MAX(0,H57-J57)","""0"""
Calculation,L57,"=IFERROR(IF(A57>n,0,IF(A57<t,H57/G57,SumInsured)),0)","""0"""
Calculation,A58,42,42.0
Calculation,B58,"=IF(A58<=n,act_nGrAx(x+$A58,MAX(0,n-$A58),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A58,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C58,"=Act_axn_k(x+$A58,MAX(0,n-$A58),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D58,"=Act_axn_k(x+$A58,MAX(0,t-$A58),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E58,"=B58-P_xt*D58+gamma2*(C58-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D58)",0.0
Calculation,F58,=SumInsured*E58,"""0"""
Calculation,G58,=B58+gamma3*C58,0.0
Calculation,H58,"=F58+alpha*t*GrossAnnualPrem*Act_axn_k(x+A58,MAX(5-A58,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I58,"=IF(AND(x+A58>=MinAgeFlex,A58>=n-MinTermFlex),1,0)",1.0
Calculation,J58,"=IF(OR(A58>n,I58),0,MIN(150,MAX(50,1%*(SumInsured-F58))))","""0"""
Calculation,K58,"=MAX(0,H58-J58)","""0"""
Calculation,L58,"=IFERROR(IF(A58>n,0,IF(A58<t,H58/G58,SumInsured)),0)","""0"""
Calculation,A59,43,43.0
Calculation,B59,"=IF(A59<=n,act_nGrAx(x+$A59,MAX(0,n-$A59),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A59,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C59,"=Act_axn_k(x+$A59,MAX(0,n-$A59),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D59,"=Act_axn_k(x+$A59,MAX(0,t-$A59),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E59,"=B59-P_xt*D59+gamma2*(C59-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D59)",0.0
Calculation,F59,=SumInsured*E59,"""0"""
Calculation,G59,=B59+gamma3*C59,0.0
Calculation,H59,"=F59+alpha*t*GrossAnnualPrem*Act_axn_k(x+A59,MAX(5-A59,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I59,"=IF(AND(x+A59>=MinAgeFlex,A59>=n-MinTermFlex),1,0)",1.0
Calculation,J59,"=IF(OR(A59>n,I59),0,MIN(150,MAX(50,1%*(SumInsured-F59))))","""0"""
Calculation,K59,"=MAX(0,H59-J59)","""0"""
Calculation,L59,"=IFERROR(IF(A59>n,0,IF(A59<t,H59/G59,SumInsured)),0)","""0"""
Calculation,A60,44,44.0
Calculation,B60,"=IF(A60<=n,act_nGrAx(x+$A60,MAX(0,n-$A60),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A60,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C60,"=Act_axn_k(x+$A60,MAX(0,n-$A60),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D60,"=Act_axn_k(x+$A60,MAX(0,t-$A60),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E60,"=B60-P_xt*D60+gamma2*(C60-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D60)",0.0
Calculation,F60,=SumInsured*E60,"""0"""
Calculation,G60,=B60+gamma3*C60,0.0
Calculation,H60,"=F60+alpha*t*GrossAnnualPrem*Act_axn_k(x+A60,MAX(5-A60,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I60,"=IF(AND(x+A60>=MinAgeFlex,A60>=n-MinTermFlex),1,0)",1.0
Calculation,J60,"=IF(OR(A60>n,I60),0,MIN(150,MAX(50,1%*(SumInsured-F60))))","""0"""
Calculation,K60,"=MAX(0,H60-J60)","""0"""
Calculation,L60,"=IFERROR(IF(A60>n,0,IF(A60<t,H60/G60,SumInsured)),0)","""0"""
Calculation,A61,45,45.0
Calculation,B61,"=IF(A61<=n,act_nGrAx(x+$A61,MAX(0,n-$A61),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A61,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C61,"=Act_axn_k(x+$A61,MAX(0,n-$A61),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D61,"=Act_axn_k(x+$A61,MAX(0,t-$A61),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E61,"=B61-P_xt*D61+gamma2*(C61-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D61)",0.0
Calculation,F61,=SumInsured*E61,"""0"""
Calculation,G61,=B61+gamma3*C61,0.0
Calculation,H61,"=F61+alpha*t*GrossAnnualPrem*Act_axn_k(x+A61,MAX(5-A61,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I61,"=IF(AND(x+A61>=MinAgeFlex,A61>=n-MinTermFlex),1,0)",1.0
Calculation,J61,"=IF(OR(A61>n,I61),0,MIN(150,MAX(50,1%*(SumInsured-F61))))","""0"""
Calculation,K61,"=MAX(0,H61-J61)","""0"""
Calculation,L61,"=IFERROR(IF(A61>n,0,IF(A61<t,H61/G61,SumInsured)),0)","""0"""
Calculation,A62,46,46.0
Calculation,B62,"=IF(A62<=n,act_nGrAx(x+$A62,MAX(0,n-$A62),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A62,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C62,"=Act_axn_k(x+$A62,MAX(0,n-$A62),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D62,"=Act_axn_k(x+$A62,MAX(0,t-$A62),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E62,"=B62-P_xt*D62+gamma2*(C62-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D62)",0.0
Calculation,F62,=SumInsured*E62,"""0"""
Calculation,G62,=B62+gamma3*C62,0.0
Calculation,H62,"=F62+alpha*t*GrossAnnualPrem*Act_axn_k(x+A62,MAX(5-A62,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I62,"=IF(AND(x+A62>=MinAgeFlex,A62>=n-MinTermFlex),1,0)",1.0
Calculation,J62,"=IF(OR(A62>n,I62),0,MIN(150,MAX(50,1%*(SumInsured-F62))))","""0"""
Calculation,K62,"=MAX(0,H62-J62)","""0"""
Calculation,L62,"=IFERROR(IF(A62>n,0,IF(A62<t,H62/G62,SumInsured)),0)","""0"""
Calculation,A63,47,47.0
Calculation,B63,"=IF(A63<=n,act_nGrAx(x+$A63,MAX(0,n-$A63),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A63,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C63,"=Act_axn_k(x+$A63,MAX(0,n-$A63),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D63,"=Act_axn_k(x+$A63,MAX(0,t-$A63),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E63,"=B63-P_xt*D63+gamma2*(C63-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D63)",0.0
Calculation,F63,=SumInsured*E63,"""0"""
Calculation,G63,=B63+gamma3*C63,0.0
Calculation,H63,"=F63+alpha*t*GrossAnnualPrem*Act_axn_k(x+A63,MAX(5-A63,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I63,"=IF(AND(x+A63>=MinAgeFlex,A63>=n-MinTermFlex),1,0)",1.0
Calculation,J63,"=IF(OR(A63>n,I63),0,MIN(150,MAX(50,1%*(SumInsured-F63))))","""0"""
Calculation,K63,"=MAX(0,H63-J63)","""0"""
Calculation,L63,"=IFERROR(IF(A63>n,0,IF(A63<t,H63/G63,SumInsured)),0)","""0"""
Calculation,A64,48,48.0
Calculation,B64,"=IF(A64<=n,act_nGrAx(x+$A64,MAX(0,n-$A64),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A64,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C64,"=Act_axn_k(x+$A64,MAX(0,n-$A64),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D64,"=Act_axn_k(x+$A64,MAX(0,t-$A64),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E64,"=B64-P_xt*D64+gamma2*(C64-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D64)",0.0
Calculation,F64,=SumInsured*E64,"""0"""
Calculation,G64,=B64+gamma3*C64,0.0
Calculation,H64,"=F64+alpha*t*GrossAnnualPrem*Act_axn_k(x+A64,MAX(5-A64,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I64,"=IF(AND(x+A64>=MinAgeFlex,A64>=n-MinTermFlex),1,0)",1.0
Calculation,J64,"=IF(OR(A64>n,I64),0,MIN(150,MAX(50,1%*(SumInsured-F64))))","""0"""
Calculation,K64,"=MAX(0,H64-J64)","""0"""
Calculation,L64,"=IFERROR(IF(A64>n,0,IF(A64<t,H64/G64,SumInsured)),0)","""0"""
Calculation,A65,49,49.0
Calculation,B65,"=IF(A65<=n,act_nGrAx(x+$A65,MAX(0,n-$A65),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A65,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C65,"=Act_axn_k(x+$A65,MAX(0,n-$A65),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D65,"=Act_axn_k(x+$A65,MAX(0,t-$A65),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E65,"=B65-P_xt*D65+gamma2*(C65-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D65)",0.0
Calculation,F65,=SumInsured*E65,"""0"""
Calculation,G65,=B65+gamma3*C65,0.0
Calculation,H65,"=F65+alpha*t*GrossAnnualPrem*Act_axn_k(x+A65,MAX(5-A65,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I65,"=IF(AND(x+A65>=MinAgeFlex,A65>=n-MinTermFlex),1,0)",1.0
Calculation,J65,"=IF(OR(A65>n,I65),0,MIN(150,MAX(50,1%*(SumInsured-F65))))","""0"""
Calculation,K65,"=MAX(0,H65-J65)","""0"""
Calculation,L65,"=IFERROR(IF(A65>n,0,IF(A65<t,H65/G65,SumInsured)),0)","""0"""
Calculation,A66,50,50.0
Calculation,B66,"=IF(A66<=n,act_nGrAx(x+$A66,MAX(0,n-$A66),Sex,MortalityTable,InterestRate)+Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x+$A66,Sex,MortalityTable,InterestRate),0)",0.0
Calculation,C66,"=Act_axn_k(x+$A66,MAX(0,n-$A66),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,D66,"=Act_axn_k(x+$A66,MAX(0,t-$A66),Sex,MortalityTable,InterestRate,1)",0.0
Calculation,E66,"=B66-P_xt*D66+gamma2*(C66-Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)*D66)",0.0
Calculation,F66,=SumInsured*E66,"""0"""
Calculation,G66,=B66+gamma3*C66,0.0
Calculation,H66,"=F66+alpha*t*GrossAnnualPrem*Act_axn_k(x+A66,MAX(5-A66,0),Sex,MortalityTable,InterestRate,1)/Act_axn_k(x,5,Sex,MortalityTable,InterestRate,1)","""0"""
Calculation,I66,"=IF(AND(x+A66>=MinAgeFlex,A66>=n-MinTermFlex),1,0)",1.0
Calculation,J66,"=IF(OR(A66>n,I66),0,MIN(150,MAX(50,1%*(SumInsured-F66))))","""0"""
Calculation,K66,"=MAX(0,H66-J66)","""0"""
Calculation,L66,"=IFERROR(IF(A66>n,0,IF(A66<t,H66/G66,SumInsured)),0)","""0"""
MortalityTables,A3,x/y,x/y
MortalityTables,B3,DAV1994_T_M,DAV1994_T_M
MortalityTables,C3,DAV1994_T_F,DAV1994_T_F
MortalityTables,D3,DAV2008_T_M,DAV2008_T_M
MortalityTables,E3,DAV2008_T_F,DAV2008_T_F
MortalityTables,A4,0,0.0
MortalityTables,B4,0.011687,0.011687
MortalityTables,C4,0.009003,0.009003
MortalityTables,D4,0.006113,0.006113
MortalityTables,E4,0.005088,0.005088
MortalityTables,A5,1,1.0
MortalityTables,B5,0.001008,0.001008
MortalityTables,C5,0.000867,0.000867
MortalityTables,D5,0.000423,0.000423
MortalityTables,E5,0.000387,0.000387
MortalityTables,A6,2,2.0
MortalityTables,B6,0.000728,0.000728
MortalityTables,C6,0.000624,0.000624
MortalityTables,D6,0.000343,0.000343
MortalityTables,E6,0.000318,0.000318
MortalityTables,A7,3,3.0
MortalityTables,B7,0.000542,0.000542
MortalityTables,C7,0.000444,0.000444
MortalityTables,D7,0.000275,0.000275
MortalityTables,E7,0.000255,0.000255
MortalityTables,A8,4,4.0
MortalityTables,B8,0.000473,0.000473
MortalityTables,C8,0.000345,0.000345
MortalityTables,D8,0.00022,0.00022
MortalityTables,E8,0.000202,0.000202
MortalityTables,A9,5,5.0
MortalityTables,B9,0.000452,0.000452
MortalityTables,C9,0.000307,0.000307
MortalityTables,D9,0.000182,0.000182
MortalityTables,E9,0.000163,0.000163
MortalityTables,A10,6,6.0
MortalityTables,B10,0.000433,0.000433
MortalityTables,C10,0.000293,0.000293
MortalityTables,D10,0.000155,0.000155
MortalityTables,E10,0.000134,0.000134
MortalityTables,A11,7,7.0
MortalityTables,B11,0.000408,0.000408
MortalityTables,C11,0.000283,0.000283
MortalityTables,D11,0.000139,0.000139
MortalityTables,E11,0.000115,0.000115
MortalityTables,A12,8,8.0
MortalityTables,B12,0.000379,0.000379
MortalityTables,C12,0.000275,0.000275
MortalityTables,D12,0.000129,0.000129
MortalityTables,E12,0.000105,0.000105
MortalityTables,A13,9,9.0
MortalityTables,B13,0.000352,0.000352
MortalityTables,C13,0.000268,0.000268
MortalityTables,D13,0.000125,0.000125
MortalityTables,E13,0.000099,9.9e-05
MortalityTables,A14,10,10.0
MortalityTables,B14,0.000334,0.000334
MortalityTables,C14,0.000261,0.000261
MortalityTables,D14,0.000129,0.000129
MortalityTables,E14,0.000102,0.000102
MortalityTables,A15,11,11.0
MortalityTables,B15,0.000331,0.000331
MortalityTables,C15,0.00026,0.00026
MortalityTables,D15,0.000143,0.000143
MortalityTables,E15,0.000111,0.000111
MortalityTables,A16,12,12.0
MortalityTables,B16,0.00034,0.00034
MortalityTables,C16,0.000267,0.000267
MortalityTables,D16,0.000173,0.000173
MortalityTables,E16,0.000127,0.000127
MortalityTables,A17,13,13.0
MortalityTables,B17,0.000371,0.000371
MortalityTables,C17,0.000281,0.000281
MortalityTables,D17,0.000222,0.000222
MortalityTables,E17,0.000153,0.000153
MortalityTables,A18,14,14.0
MortalityTables,B18,0.000451,0.000451
MortalityTables,C18,0.000307,0.000307
MortalityTables,D18,0.000303,0.000303
MortalityTables,E18,0.000188,0.000188
MortalityTables,A19,15,15.0
MortalityTables,B19,0.000593,0.000593
MortalityTables,C19,0.000353,0.000353
MortalityTables,D19,0.000417,0.000417
MortalityTables,E19,0.000228,0.000228
MortalityTables,A20,16,16.0
MortalityTables,B20,0.000792,0.000792
MortalityTables,C20,0.000416,0.000416
MortalityTables,D20,0.000557,0.000557
MortalityTables,E20,0.000271,0.000271
MortalityTables,A21,17,17.0
MortalityTables,B21,0.00104,0.00104
MortalityTables,C21,0.00048,0.00048
MortalityTables,D21,0.000709,0.000709
MortalityTables,E21,0.00031,0.00031
MortalityTables,A22,18,18.0
MortalityTables,B22,0.001298,0.001298
MortalityTables,C22,0.000537,0.000537
MortalityTables,D22,0.00085,0.00085
MortalityTables,E22,0.000324,0.000324
MortalityTables,A23,19,19.0
MortalityTables,B23,0.001437,0.001437
MortalityTables,C23,0.00056,0.00056
MortalityTables,D23,0.000953,0.000953
MortalityTables,E23,0.00033,0.00033
MortalityTables,A24,20,20.0
MortalityTables,B24,0.001476,0.001476
MortalityTables,C24,0.00056,0.00056
MortalityTables,D24,0.001012,0.001012
MortalityTables,E24,0.000328,0.000328
MortalityTables,A25,21,21.0
MortalityTables,B25,0.001476,0.001476
MortalityTables,C25,0.00056,0.00056
MortalityTables,D25,0.001022,0.001022
MortalityTables,E25,0.000322,0.000322
MortalityTables,A26,22,22.0
MortalityTables,B26,0.001476,0.001476
MortalityTables,C26,0.00056,0.00056
MortalityTables,D26,0.001004,0.001004
MortalityTables,E26,0.000314,0.000314
MortalityTables,A27,23,23.0
MortalityTables,B27,0.001476,0.001476
MortalityTables,C27,0.00056,0.00056
MortalityTables,D27,0.000963,0.000963
MortalityTables,E27,0.000304,0.000304
MortalityTables,A28,24,24.0
MortalityTables,B28,0.001476,0.001476
MortalityTables,C28,0.00056,0.00056
MortalityTables,D28,0.000911,0.000911
MortalityTables,E28,0.000297,0.000297
MortalityTables,A29,25,25.0
MortalityTables,B29,0.001476,0.001476
MortalityTables,C29,0.00056,0.00056
MortalityTables,D29,0.000856,0.000856
MortalityTables,E29,0.000293,0.000293
MortalityTables,A30,26,26.0
MortalityTables,B30,0.001476,0.001476
MortalityTables,C30,0.00056,0.00056
MortalityTables,D30,0.000808,0.000808
MortalityTables,E30,0.000292,0.000292
MortalityTables,A31,27,27.0
MortalityTables,B31,0.001476,0.001476
MortalityTables,C31,0.000581,0.000581
MortalityTables,D31,0.000772,0.000772
MortalityTables,E31,0.000292,0.000292
MortalityTables,A32,28,28.0
MortalityTables,B32,0.001476,0.001476
MortalityTables,C32,0.000612,0.000612
MortalityTables,D32,0.000752,0.000752
MortalityTables,E32,0.000296,0.000296
MortalityTables,A33,29,29.0
MortalityTables,B33,0.001476,0.001476
MortalityTables,C33,0.000645,0.000645
MortalityTables,D33,0.000745,0.000745
MortalityTables,E33,0.000302,0.000302
MortalityTables,A34,30,30.0
MortalityTables,B34,0.001476,0.001476
MortalityTables,C34,0.000689,0.000689
MortalityTables,D34,0.000752,0.000752
MortalityTables,E34,0.000311,0.000311
MortalityTables,A35,31,31.0
MortalityTables,B35,0.001476,0.001476
MortalityTables,C35,0.000735,0.000735
MortalityTables,D35,0.000768,0.000768
MortalityTables,E35,0.000327,0.000327
MortalityTables,A36,32,32.0
MortalityTables,B36,0.001489,0.001489
MortalityTables,C36,0.000783,0.000783
MortalityTables,D36,0.000791,0.000791
MortalityTables,E36,0.000351,0.000351
MortalityTables,A37,33,33.0
MortalityTables,B37,0.001551,0.001551
MortalityTables,C37,0.000833,0.000833
MortalityTables,D37,0.00082,0.00082
MortalityTables,E37,0.000386,0.000386
MortalityTables,A38,34,34.0
MortalityTables,B38,0.001641,0.001641
MortalityTables,C38,0.000897,0.000897
MortalityTables,D38,0.000855,0.000855
MortalityTables,E38,0.000433,0.000433
MortalityTables,A39,35,35.0
MortalityTables,B39,0.001747,0.001747
MortalityTables,C39,0.000971,0.000971
MortalityTables,D39,0.000895,0.000895
MortalityTables,E39,0.00049,0.00049
MortalityTables,A40,36,36.0
MortalityTables,B40,0.001869,0.001869
MortalityTables,C40,0.001057,0.001057
MortalityTables,D40,0.000945,0.000945
MortalityTables,E40,0.000555,0.000555
MortalityTables,A41,37,37.0
MortalityTables,B41,0.002007,0.002007
MortalityTables,C41,0.001156,0.001156
MortalityTables,D41,0.001005,0.001005
MortalityTables,E41,0.000624,0.000624
MortalityTables,A42,38,38.0
MortalityTables,B42,0.002167,0.002167
MortalityTables,C42,0.001267,0.001267
MortalityTables,D42,0.001083,0.001083
MortalityTables,E42,0.000701,0.000701
MortalityTables,A43,39,39.0
MortalityTables,B43,0.002354,0.002354
MortalityTables,C43,0.00139,0.00139
MortalityTables,D43,0.001181,0.001181
MortalityTables,E43,0.000783,0.000783
MortalityTables,A44,40,40.0
MortalityTables,B44,0.002569,0.002569
MortalityTables,C44,0.001524,0.001524
MortalityTables,D44,0.001301,0.001301
MortalityTables,E44,0.000872,0.000872
MortalityTables,A45,41,41.0
MortalityTables,B45,0.002823,0.002823
MortalityTables,C45,0.001672,0.001672
MortalityTables,D45,0.001447,0.001447
MortalityTables,E45,0.000972,0.000972
MortalityTables,A46,42,42.0
MortalityTables,B46,0.003087,0.003087
MortalityTables,C46,0.001812,0.001812
MortalityTables,D46,0.001623,0.001623
MortalityTables,E46,0.001084,0.001084
MortalityTables,A47,43,43.0
MortalityTables,B47,0.003387,0.003387
MortalityTables,C47,0.001964,0.001964
MortalityTables,D47,0.001833,0.001833
MortalityTables,E47,0.001213,0.001213
MortalityTables,A48,44,44.0
MortalityTables,B48,0.003726,0.003726
MortalityTables,C48,0.002126,0.002126
MortalityTables,D48,0.002082,0.002082
MortalityTables,E48,0.001359,0.001359
MortalityTables,A49,45,45.0
MortalityTables,B49,0.0041,0.0041
MortalityTables,C49,0.002295,0.002295
MortalityTables,D49,0.002364,0.002364
MortalityTables,E49,0.001524,0.001524
MortalityTables,A50,46,46.0
MortalityTables,B50,0.004522,0.004522
MortalityTables,C50,0.00248,0.00248
MortalityTables,D50,0.002669,0.002669
MortalityTables,E50,0.001706,0.001706
MortalityTables,A51,47,47.0
MortalityTables,B51,0.004983,0.004983
MortalityTables,C51,0.002676,0.002676
MortalityTables,D51,0.002983,0.002983
MortalityTables,E51,0.001903,0.001903
MortalityTables,A52,48,48.0
MortalityTables,B52,0.005508,0.005508
MortalityTables,C52,0.002902,0.002902
MortalityTables,D52,0.003302,0.003302
MortalityTables,E52,0.002109,0.002109
MortalityTables,A53,49,49.0
MortalityTables,B53,0.006094,0.006094
MortalityTables,C53,0.003151,0.003151
MortalityTables,D53,0.00363,0.00363
MortalityTables,E53,0.002324,0.002324
MortalityTables,A54,50,50.0
MortalityTables,B54,0.006751,0.006751
MortalityTables,C54,0.003425,0.003425
MortalityTables,D54,0.003981,0.003981
MortalityTables,E54,0.002546,0.002546
MortalityTables,A55,51,51.0
MortalityTables,B55,0.007485,0.007485
MortalityTables,C55,0.003728,0.003728
MortalityTables,D55,0.004371,0.004371
MortalityTables,E55,0.002782,0.002782
MortalityTables,A56,52,52.0
MortalityTables,B56,0.008302,0.008302
MortalityTables,C56,0.004066,0.004066
MortalityTables,D56,0.004812,0.004812
MortalityTables,E56,0.003035,0.003035
MortalityTables,A57,53,53.0
MortalityTables,B57,0.009215,0.009215
MortalityTables,C57,0.00445,0.00445
MortalityTables,D57,0.005308,0.005308
MortalityTables,E57,0.003306,0.003306
MortalityTables,A58,54,54.0
MortalityTables,B58,0.010195,0.010195
MortalityTables,C58,0.004862,0.004862
MortalityTables,D58,0.005857,0.005857
MortalityTables,E58,0.003593,0.003593
MortalityTables,A59,55,55.0
MortalityTables,B59,0.011236,0.011236
MortalityTables,C59,0.005303,0.005303
MortalityTables,D59,0.00646,0.00646
MortalityTables,E59,0.003898,0.003898
MortalityTables,A60,56,56.0
MortalityTables,B60,0.01234,0.01234
MortalityTables,C60,0.005777,0.005777
MortalityTables,D60,0.007117,0.007117
MortalityTables,E60,0.004228,0.004228
MortalityTables,A61,57,57.0
MortalityTables,B61,0.013519,0.013519
MortalityTables,C61,0.006302,0.006302
MortalityTables,D61,0.007831,0.007831
MortalityTables,E61,0.004585,0.004585
MortalityTables,A62,58,58.0
MortalityTables,B62,0.014784,0.014784
MortalityTables,C62,0.006884,0.006884
MortalityTables,D62,0.008604,0.008604
MortalityTables,E62,0.004974,0.004974
MortalityTables,A63,59,59.0
MortalityTables,B63,0.01615,0.01615
MortalityTables,C63,0.00753,0.00753
MortalityTables,D63,0.009454,0.009454
MortalityTables,E63,0.005402,0.005402
MortalityTables,A64,60,60.0
MortalityTables,B64,0.017625,0.017625
MortalityTables,C64,0.00824,0.00824
MortalityTables,D64,0.010404,0.010404
MortalityTables,E64,0.005884,0.005884
MortalityTables,A65,61,61.0
MortalityTables,B65,0.019223,0.019223
MortalityTables,C65,0.009022,0.009022
MortalityTables,D65,0.011504,0.011504
MortalityTables,E65,0.006449,0.006449
MortalityTables,A66,62,62.0
MortalityTables,B66,0.020956,0.020956
MortalityTables,C66,0.009884,0.009884
MortalityTables,D66,0.012818,0.012818
MortalityTables,E66,0.007126,0.007126
MortalityTables,A67,63,63.0
MortalityTables,B67,0.022833,0.022833
MortalityTables,C67,0.010839,0.010839
MortalityTables,D67,0.014429,0.014429
MortalityTables,E67,0.007935,0.007935
MortalityTables,A68,64,64.0
MortalityTables,B68,0.024858,0.024858
MortalityTables,C68,0.011889,0.011889
MortalityTables,D68,0.016415,0.016415
MortalityTables,E68,0.008898,0.008898
MortalityTables,A69,65,65.0
MortalityTables,B69,0.027073,0.027073
MortalityTables,C69,0.013054,0.013054
MortalityTables,D69,0.018832,0.018832
MortalityTables,E69,0.010025,0.010025
MortalityTables,A70,66,66.0
MortalityTables,B70,0.029552,0.029552
MortalityTables,C70,0.014371,0.014371
MortalityTables,D70,0.021704,0.021704
MortalityTables,E70,0.011323,0.011323
MortalityTables,A71,67,67.0
MortalityTables,B71,0.03235,0.03235
MortalityTables,C71,0.015874,0.015874
MortalityTables,D71,0.025016,0.025016
MortalityTables,E71,0.012797,0.012797
MortalityTables,A72,68,68.0
MortalityTables,B72,0.035632,0.035632
MortalityTables,C72,0.017667,0.017667
MortalityTables,D72,0.028738,0.028738
MortalityTables,E72,0.01446,0.01446
MortalityTables,A73,69,69.0
MortalityTables,B73,0.039224,0.039224
MortalityTables,C73,0.019657,0.019657
MortalityTables,D73,0.032822,0.032822
MortalityTables,E73,0.016332,0.016332
MortalityTables,A74,70,70.0
MortalityTables,B74,0.043127,0.043127
MortalityTables,C74,0.021861,0.021861
MortalityTables,D74,0.037219,0.037219
MortalityTables,E74,0.01844,0.01844
MortalityTables,A75,71,71.0
MortalityTables,B75,0.0474,0.0474
MortalityTables,C75,0.024344,0.024344
MortalityTables,D75,0.04188,0.04188
MortalityTables,E75,0.020813,0.020813
MortalityTables,A76,72,72.0
MortalityTables,B76,0.05211,0.05211
MortalityTables,C76,0.027191,0.027191
MortalityTables,D76,0.046597,0.046597
MortalityTables,E76,0.023475,0.023475
MortalityTables,A77,73,73.0
MortalityTables,B77,0.057472,0.057472
MortalityTables,C77,0.030576,0.030576
MortalityTables,D77,0.051181,0.051181
MortalityTables,E77,0.027035,0.027035
MortalityTables,A78,74,74.0
MortalityTables,B78,0.06344,0.06344
MortalityTables,C78,0.034504,0.034504
MortalityTables,D78,0.05611,0.05611
MortalityTables,E78,0.030413,0.030413
MortalityTables,A79,75,75.0
MortalityTables,B79,0.070039,0.070039
MortalityTables,C79,0.03903,0.03903
MortalityTables,D79,0.061477,0.061477
MortalityTables,E79,0.034287,0.034287
MortalityTables,A80,76,76.0
MortalityTables,B80,0.077248,0.077248
MortalityTables,C80,0.044184,0.044184
MortalityTables,D80,0.067433,0.067433
MortalityTables,E80,0.038749,0.038749
MortalityTables,A81,77,77.0
MortalityTables,B81,0.085073,0.085073
MortalityTables,C81,0.050014,0.050014
MortalityTables,D81,0.07416,0.07416
MortalityTables,E81,0.043937,0.043937
MortalityTables,A82,78,78.0
MortalityTables,B82,0.093534,0.093534
MortalityTables,C82,0.056574,0.056574
MortalityTables,D82,0.081806,0.081806
MortalityTables,E82,0.049993,0.049993
MortalityTables,A83,79,79.0
MortalityTables,B83,0.102662,0.102662
MortalityTables,C83,0.063921,0.063921
MortalityTables,D83,0.090478,0.090478
MortalityTables,E83,0.057024,0.057024
MortalityTables,A84,80,80.0
MortalityTables,B84,0.112477,0.112477
MortalityTables,C84,0.072101,0.072101
MortalityTables,D84,0.100261,0.100261
MortalityTables,E84,0.065113,0.065113
MortalityTables,A85,81,81.0
MortalityTables,B85,0.122995,0.122995
MortalityTables,C85,0.081151,0.081151
MortalityTables,D85,0.111193,0.111193
MortalityTables,E85,0.074288,0.074288
MortalityTables,A86,82,82.0
MortalityTables,B86,0.134231,0.134231
MortalityTables,C86,0.091096,0.091096
MortalityTables,D86,0.123283,0.123283
MortalityTables,E86,0.08459,0.08459
MortalityTables,A87,83,83.0
MortalityTables,B87,0.146212,0.146212
MortalityTables,C87,0.10197,0.10197
MortalityTables,D87,0.136498,0.136498
MortalityTables,E87,0.096095,0.096095
MortalityTables,A88,84,84.0
MortalityTables,B88,0.158964,0.158964
MortalityTables,C88,0.113798,0.113798
MortalityTables,D88,0.150887,0.150887
MortalityTables,E88,0.109028,0.109028
MortalityTables,A89,85,85.0
MortalityTables,B89,0.172512,0.172512
MortalityTables,C89,0.126628,0.126628
MortalityTables,D89,0.1665,0.1665
MortalityTables,E89,0.123611,0.123611
MortalityTables,A90,86,86.0
MortalityTables,B90,0.186896,0.186896
MortalityTables,C90,0.140479,0.140479
MortalityTables,D90,0.183344,0.183344
MortalityTables,E90,0.140022,0.140022
MortalityTables,A91,87,87.0
MortalityTables,B91,0.202185,0.202185
MortalityTables,C91,0.155379,0.155379
MortalityTables,D91,0.201323,0.201323
MortalityTables,E91,0.158257,0.158257
MortalityTables,A92,88,88.0
MortalityTables,B92,0.218413,0.218413
MortalityTables,C92,0.171325,0.171325
MortalityTables,D92,0.220284,0.220284
MortalityTables,E92,0.178185,0.178185
MortalityTables,A93,89,89.0
MortalityTables,B93,0.235597,0.235597
MortalityTables,C93,0.188318,0.188318
MortalityTables,D93,0.240073,0.240073
MortalityTables,E93,0.199669,0.199669
MortalityTables,A94,90,90.0
MortalityTables,B94,0.253691,0.253691
MortalityTables,C94,0.206375,0.206375
MortalityTables,D94,0.260556,0.260556
MortalityTables,E94,0.222504,0.222504
MortalityTables,A95,91,91.0
MortalityTables,B95,0.272891,0.272891
MortalityTables,C95,0.225558,0.225558
MortalityTables,D95,0.281602,0.281602
MortalityTables,E95,0.246453,0.246453
MortalityTables,A96,92,92.0
MortalityTables,B96,0.293142,0.293142
MortalityTables,C96,0.245839,0.245839
MortalityTables,D96,0.303079,0.303079
MortalityTables,E96,0.271195,0.271195
MortalityTables,A97,93,93.0
MortalityTables,B97,0.314638,0.314638
MortalityTables,C97,0.26727,0.26727
MortalityTables,D97,0.324872,0.324872
MortalityTables,E97,0.295584,0.295584
MortalityTables,A98,94,94.0
MortalityTables,B98,0.337739,0.337739
MortalityTables,C98,0.289983,0.289983
MortalityTables,D98,0.346887,0.346887
MortalityTables,E98,0.319362,0.319362
MortalityTables,A99,95,95.0
MortalityTables,B99,0.36206,0.36206
MortalityTables,C99,0.314007,0.314007
MortalityTables,D99,0.369051,0.369051
MortalityTables,E99,0.343441,0.343441
MortalityTables,A100,96,96.0
MortalityTables,B100,0.388732,0.388732
MortalityTables,C100,0.340119,0.340119
MortalityTables,D100,0.391305,0.391305
MortalityTables,E100,0.367818,0.367818
MortalityTables,A101,97,97.0
MortalityTables,B101,0.419166,0.419166
MortalityTables,C101,0.367388,0.367388
MortalityTables,D101,0.413938,0.413938
MortalityTables,E101,0.392493,0.392493
MortalityTables,A102,98,98.0
MortalityTables,B102,0.452008,0.452008
MortalityTables,C102,0.397027,0.397027
MortalityTables,D102,0.437313,0.437313
MortalityTables,E102,0.41746,0.41746
MortalityTables,A103,99,99.0
MortalityTables,B103,0.4864,0.4864
MortalityTables,C103,0.428748,0.428748
MortalityTables,D103,0.461101,0.461101
MortalityTables,E103,0.442716,0.442716
MortalityTables,A104,100,100.0
MortalityTables,B104,1,1.0
MortalityTables,C104,1,1.0
MortalityTables,D104,0.485304,0.485304
MortalityTables,E104,0.468258,0.468258
MortalityTables,A105,101,101.0
MortalityTables,B105,1,1.0
MortalityTables,C105,1,1.0
MortalityTables,D105,0.509924,0.509924
MortalityTables,E105,0.494075,0.494075
MortalityTables,A106,102,102.0
MortalityTables,B106,1,1.0
MortalityTables,C106,1,1.0
MortalityTables,D106,0.534957,0.534957
MortalityTables,E106,0.520164,0.520164
MortalityTables,A107,103,103.0
MortalityTables,B107,1,1.0
MortalityTables,C107,1,1.0
MortalityTables,D107,0.560407,0.560407
MortalityTables,E107,0.546514,0.546514
MortalityTables,A108,104,104.0
MortalityTables,B108,1,1.0
MortalityTables,C108,1,1.0
MortalityTables,D108,0.586265,0.586265
MortalityTables,E108,0.573114,0.573114
MortalityTables,A109,105,105.0
MortalityTables,B109,1,1.0
MortalityTables,C109,1,1.0
MortalityTables,D109,0.612529,0.612529
MortalityTables,E109,0.599953,0.599953
MortalityTables,A110,106,106.0
MortalityTables,B110,1,1.0
MortalityTables,C110,1,1.0
MortalityTables,D110,0.639188,0.639188
MortalityTables,E110,0.627014,0.627014
MortalityTables,A111,107,107.0
MortalityTables,B111,1,1.0
MortalityTables,C111,1,1.0
MortalityTables,D111,0.666233,0.666233
MortalityTables,E111,0.654283,0.654283
MortalityTables,A112,108,108.0
MortalityTables,B112,1,1.0
MortalityTables,C112,1,1.0
MortalityTables,D112,0.693651,0.693651
MortalityTables,E112,0.681741,0.681741
MortalityTables,A113,109,109.0
MortalityTables,B113,1,1.0
MortalityTables,C113,1,1.0
MortalityTables,D113,0.721425,0.721425
MortalityTables,E113,0.709364,0.709364
MortalityTables,A114,110,110.0
MortalityTables,B114,1,1.0
MortalityTables,C114,1,1.0
MortalityTables,D114,0.749533,0.749533
MortalityTables,E114,0.73713,0.73713
MortalityTables,A115,111,111.0
MortalityTables,B115,1,1.0
MortalityTables,C115,1,1.0
MortalityTables,D115,0.77795,0.77795
MortalityTables,E115,0.765011,0.765011
MortalityTables,A116,112,112.0
MortalityTables,B116,1,1.0
MortalityTables,C116,1,1.0
MortalityTables,D116,0.806647,0.806647
MortalityTables,E116,0.792974,0.792974
MortalityTables,A117,113,113.0
MortalityTables,B117,1,1.0
MortalityTables,C117,1,1.0
MortalityTables,D117,0.835585,0.835585
MortalityTables,E117,0.820987,0.820987
MortalityTables,A118,114,114.0
MortalityTables,B118,1,1.0
MortalityTables,C118,1,1.0
MortalityTables,D118,0.864722,0.864722
MortalityTables,E118,0.849009,0.849009
MortalityTables,A119,115,115.0
MortalityTables,B119,1,1.0
MortalityTables,C119,1,1.0
MortalityTables,D119,0.894008,0.894008
MortalityTables,E119,0.876998,0.876998
MortalityTables,A120,116,116.0
MortalityTables,B120,1,1.0
MortalityTables,C120,1,1.0
MortalityTables,D120,0.923382,0.923382
MortalityTables,E120,0.904905,0.904905
MortalityTables,A121,117,117.0
MortalityTables,B121,1,1.0
MortalityTables,C121,1,1.0
MortalityTables,D121,0.952778,0.952778
MortalityTables,E121,0.932675,0.932675
MortalityTables,A122,118,118.0
MortalityTables,B122,1,1.0
MortalityTables,C122,1,1.0
MortalityTables,D122,0.982113,0.982113
MortalityTables,E122,0.960249,0.960249
MortalityTables,A123,119,119.0
MortalityTables,B123,1,1.0
MortalityTables,C123,1,1.0
MortalityTables,D123,1,1.0
MortalityTables,E123,0.987564,0.987564
MortalityTables,A124,120,120.0
MortalityTables,B124,1,1.0
MortalityTables,C124,1,1.0
MortalityTables,D124,1,1.0
MortalityTables,E124,1,1.0
MortalityTables,A125,121,121.0
MortalityTables,B125,1,1.0
MortalityTables,C125,1,1.0
MortalityTables,D125,1,1.0
MortalityTables,E125,1,1.0
MortalityTables,A126,122,122.0
MortalityTables,B126,1,1.0
MortalityTables,C126,1,1.0
MortalityTables,D126,1,1.0
MortalityTables,E126,1,1.0
MortalityTables,A127,123,123.0
MortalityTables,B127,1,1.0
MortalityTables,C127,1,1.0
MortalityTables,D127,1,1.0
MortalityTables,E127,1,1.0

---

### FILENAME: excelrange.csv
CONTENT:

Sheet,Name,Address
,_xleta.MAX,=#NAME?
,_xlfn.IFERROR,=#NAME?
Calculation,alpha,=Calculation!$E$6
Calculation,beta1,=Calculation!$E$7
Calculation,gamma1,=Calculation!$E$8
Calculation,gamma2,=Calculation!$E$9
Calculation,gamma3,=Calculation!$E$10
Calculation,GrossAnnualPrem,=Calculation!$K$6
Calculation,GrossModalPrem,=Calculation!$K$7
Calculation,InterestRate,=Calculation!$E$4
Calculation,k,=Calculation!$E$11
MortalityTables,m_Tables,=MortalityTables!$B$4:$E$127
Calculation,MinAgeFlex,=Calculation!$H$4
Calculation,MinTermFlex,=Calculation!$H$5
Calculation,ModalSurcharge,=Calculation!$E$12
Calculation,MortalityTable,=Calculation!$E$5
Calculation,n,=Calculation!$B$6
Calculation,NormGrossAnnualPrem,=Calculation!$K$5
Calculation,P_xt,=Calculation!$K$9
Calculation,PayFreq,=Calculation!$B$9
Calculation,Sex,=Calculation!$B$5
Calculation,SumInsured,=Calculation!$B$8
Calculation,t,=Calculation!$B$7
MortalityTables,v_Age,=MortalityTables!$A$4:$A$127
MortalityTables,v_Tables,=MortalityTables!$B$3:$E$3
Calculation,x,=Calculation!$B$4


---

### FILENAME: protocol.txt
CONTENT:

Du:
# Prompt Series  
**Project:** Excel Product Calculator → Python Product Calculator  

---

## GLOBAL CONTEXT

### Role and Objective  
You are a **Senior Python Engineer**.  
Goal: From the Excel file `Tariff_Calculator.xlsm`, create a **modular, pure‑Python product calculator** that produces identical results.  
The solution consists of **seven scripts** and several **.csv input files** (see table below).  
Your working directory is root.  
The Excel file is located in the `input` directory.  
All outputs generated by the LLM and by the scripts must be written to the project root directory.  


### Deliverables
| Code | File | Content | Acceptance Criteria |
|--------|-------|--------|---------------|
| EXCEL_TO_TEXT | `excel_to_text.py` | Extract cells & ranges → `excelcell.csv`, `excelrange.csv` | Both CSVs exist & have ≥ 1 row |
| VBA_TO_TEXT | `vba_to_text.py` | Extract all VBA modules → `Mod_*.txt` | All module files exist |
| DATA_EXTRACT | `data_extract.py` | Create `var.csv`, `tariff.csv`, `limits.csv`, `tables.csv`, `tariff.py` | All files exist & have ≥ 1 data row |
| BASISFUNCT | `basfunct.py` | 1‑to‑1 port of the VBA base functions | pytest suite passes |
| AUSFUNCT_T1 | `outfunc.py` | Contains `NormGrossAnnualPrem()` and dependent functions | `NormGrossAnnualPrem()` test < 1 e‑6 deviation |
| AUSFUNCT_T2 | `outfunc2.py` | Additional output functions | All function tests pass |
| CLI_RUNNER | `run_calc.py` | Command‑line interface | `python run_calc.py --help` runs |

### General Rules  
- **Language:** English in documentation, variable names in English (`present_value`).  
- **Quality:** Black formatting, Ruff lint = 0 warnings.  
- **Response format:** Each task delivers as many executable code blocks as defined in the relevant task.  
- **Progress:** Work **strictly step by step**. Complete each task, wait for my **“✅”**, only then continue. DO NOT START any coding if note explicitely instructed to!  

ChatGPT:
Understood. I’ll work step-by-step, produce exactly one executable code block per task, write all generated outputs to /Bartek project root, and not start any coding until you explicitly tell me to.

When you’re ready, tell me which deliverable to start with (recommended order is EXCEL_TO_TEXT → VBA_TO_TEXT → DATA_EXTRACT → basfunct → outfunc → outfunc2 → run_calc).


Du:
## TASK 1 – Implement Code for Context Export (Excel → CSV)

1. **Goal**: Create a simple script to tranlsate Excel-Content in text (`excel_to_text.py`)
2. **Input for the script**: `Tariff_Calculator.xlsm`  
3. **Output fo the script**  
   - `excelcell.csv` (columns: Sheet, Address, Formula, Value)  
   - `excelrange.csv` (columns: Sheet, Name, Address)  
4. **Procedure for the script**  
   - Use `xlwings` for array formulas.  
   - Ignore empty cells.  
   - Prerequisite: `xlwings` requires a local Excel installation.  
5. **Success Check**  
```python
assert Path("excelcell.csv").stat().st_size > 10_000
assert "Calculation" in pd.read_csv("excelcell.csv")["Sheet"].unique()
```  
6. **Delivery format**: Full, executable code block.

ChatGPT:
#!/usr/bin/env python3
"""
EXCEL_TO_TEXT: Export Excel workbook content to CSV.

Input:
  input/Tariff_Calculator.xlsm

Outputs (written to project root / current working directory):
  - excelcell.csv  columns: Sheet, Address, Formula, Value
  - excelrange.csv columns: Sheet, Name, Address

Notes:
  - Uses xlwings (requires local Excel installation).
  - Ignores empty cells (both formula and value empty/None).
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Any, Iterable, Optional, Tuple

import xlwings as xw


INPUT_PATH = Path("input") / "Tariff_Calculator.xlsm"
OUT_CELLS = Path("excelcell.csv")
OUT_RANGES = Path("excelrange.csv")


def col_to_letters(col_num: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA"""
    if col_num < 1:
        raise ValueError(f"Invalid column number: {col_num}")
    letters = []
    while col_num:
        col_num, rem = divmod(col_num - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def a1_address(row: int, col: int) -> str:
    return f"{col_to_letters(col)}{row}"


def stringify_value(v: Any) -> str:
    if v is None:
        return ""
    # Excel errors sometimes come through as strings like '#N/A', keep as-is
    if isinstance(v, (str, int, float, bool)):
        return str(v)
    # Datetimes, arrays, etc.
    try:
        return json.dumps(v, ensure_ascii=False, default=str)
    except TypeError:
        return str(v)


def normalize_formula(f: Any) -> str:
    if f is None:
        return ""
    if isinstance(f, str):
        return f
    # xlwings can return nested lists for multi-area; serialize to JSON for safety
    try:
        return json.dumps(f, ensure_ascii=False, default=str)
    except TypeError:
        return str(f)


_REFERS_TO_SHEET_RE = re.compile(
    r"""^=?'?(?P<sheet>[^']+?)'?!""", re.IGNORECASE
)


def parse_sheet_from_refers_to(refers_to: str) -> str:
    """
    Try to extract the sheet name from a Name.RefersTo string like:
      ="'Calculation'!$A$1:$B$2"
      ="=Calculation!$A$1"
    """
    if not refers_to:
        return ""
    s = refers_to.strip()
    if s.startswith("="):
        s = s[1:].lstrip()
    m = _REFERS_TO_SHEET_RE.match(s)
    return m.group("sheet") if m else ""


def is_empty_cell(formula: str, value: Any) -> bool:
    # Treat as empty if no formula and value is None/"".
    if formula and formula != "":
        return False
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def iter_used_range_cells(
    sheet: xw.Sheet,
) -> Iterable[Tuple[str, str, str, Any]]:
    """
    Yield (sheet_name, address, formula, value) for non-empty cells in sheet.used_range.
    Uses array reads for performance and to preserve array formulas via Excel.
    """
    used = sheet.used_range
    # If the sheet is truly empty, used_range may still return A1; we will filter empties.
    top_row = used.row
    left_col = used.column
    nrows = used.rows.count
    ncols = used.columns.count

    # Read formulas & values as 2D arrays in a single call each
    formulas = used.formula  # can be scalar if single cell
    values = used.value

    # Normalize to 2D lists
    if nrows == 1 and ncols == 1:
        formulas_2d = [[formulas]]
        values_2d = [[values]]
    else:
        # xlwings returns list-of-lists for ranges; ensure shape
        formulas_2d = formulas
        values_2d = values

    for r in range(nrows):
        row_idx = top_row + r
        for c in range(ncols):
            col_idx = left_col + c
            f = normalize_formula(formulas_2d[r][c])
            v = values_2d[r][c]
            if is_empty_cell(f, v):
                continue
            addr = a1_address(row_idx, col_idx)
            yield (sheet.name, addr, f, v)


def export_cells(book: xw.Book) -> int:
    count = 0
    with OUT_CELLS.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Sheet", "Address", "Formula", "Value"])
        for sht in book.sheets:
            for sheet_name, addr, formula, value in iter_used_range_cells(sht):
                w.writerow([sheet_name, addr, formula, stringify_value(value)])
                count += 1
    return count


def export_named_ranges(book: xw.Book) -> int:
    count = 0
    with OUT_RANGES.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Sheet", "Name", "Address"])
        for nm in book.names:
            try:
                name = nm.name
                refers_to = nm.refers_to  # e.g., ="'Sheet'!$A$1:$B$2"
            except Exception:
                # Skip names we cannot read (rare)
                continue
            sheet = parse_sheet_from_refers_to(refers_to)
            if not name or not refers_to:
                continue
            w.writerow([sheet, name, refers_to])
            count += 1
    return count


def main() -> None:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Excel input not found: {INPUT_PATH.resolve()}")

    app: Optional[xw.App] = None
    book: Optional[xw.Book] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        book = app.books.open(str(INPUT_PATH), update_links=False, read_only=True)

        cell_rows = export_cells(book)
        range_rows = export_named_ranges(book)

        # Basic sanity output for humans
        print(f"Wrote {OUT_CELLS} with {cell_rows} data rows.")
        print(f"Wrote {OUT_RANGES} with {range_rows} data rows.")
        print(f"Workbook: {INPUT_PATH}")

    finally:
        try:
            if book is not None:
                book.close()
        finally:
            if app is not None:
                app.quit()


if __name__ == "__main__":
    main()

Du:
## TASK 2 – Implement Code for VBA Export (VBA → TXT)

1. **Goal**: Create a simple script to tranlsate Excel-VBA-Module-Content in text (`vba_to_text.py`)
2. **Input for the script**: `Tariff_Calculator.xlsm`  
3. **Output fo the script**: One `Mod_*.txt` file per VBA module  
4. **Procedure for the script**  
   - Use `oletools.olevba` or `vb2py` to dump the modules.  
   - Filename scheme: `Mod_<Name>.txt`.  
   - Process all non‑empty code modules, including those without `Sub` (e. g., with constants).    
   - Ignore empty modules or code objects (e. g., Excel worksheet without code).  
5. **Success Check**  
   - Number of `.txt` files ≥ number of non‑empty modules in the VBA editor.  
   - Each file contains at least one `Sub` or `Function`.

ChatGPT:
#!/usr/bin/env python3
"""
VBA_TO_TEXT: Export all non-empty VBA modules from an .xlsm into text files.

Input:
  input/Tariff_Calculator.xlsm

Output (written to project root / current working directory):
  - Mod_<Name>.txt for each non-empty VBA code module

Rules implemented:
  - Uses oletools.olevba to extract modules.
  - Writes one file per module (deduplicates by output filename).
  - Ignores empty modules / code objects with no code text.
  - Includes modules even if they only contain constants (no Sub/Function).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict

from oletools.olevba import VBA_Parser  # type: ignore


INPUT_PATH = Path("input") / "Tariff_Calculator.xlsm"
OUT_PREFIX = "Mod_"

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]+')
SUBFUNC_RE = re.compile(r"(?im)^\s*(public|private|friend)?\s*(sub|function)\b")


def sanitize_filename(name: str) -> str:
    name = (name or "").strip()
    name = INVALID_FILENAME_CHARS.sub("_", name)
    name = re.sub(r"\s+", "_", name)
    name = name.strip("._")
    return name or "Unnamed"


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def pick_module_name(vba_filename: str, stream_path: str) -> str:
    name = (vba_filename or "").strip()
    if name:
        return name
    parts = (stream_path or "").replace("\\", "/").split("/")
    return parts[-1] if parts and parts[-1] else "Unnamed"


def unique_path(base: Path, used: set[str]) -> Path:
    """Ensure filename uniqueness within this run."""
    if base.name not in used:
        used.add(base.name)
        return base
    stem, suffix = base.stem, base.suffix
    i = 2
    while True:
        candidate = base.with_name(f"{stem}_{i}{suffix}")
        if candidate.name not in used:
            used.add(candidate.name)
            return candidate
        i += 1


def main() -> None:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Excel input not found: {INPUT_PATH.resolve()}")

    out_dir = Path(".")
    used_filenames: set[str] = set()
    written: Dict[str, Path] = {}
    exported_count = 0
    skipped_empty = 0
    no_subfunc_count = 0

    vba = VBA_Parser(str(INPUT_PATH))
    try:
        if not vba.detect_vba_macros():
            print("No VBA macros detected.")
            return

        for (_container, stream_path, vba_filename, code) in vba.extract_macros():
            if code is None:
                skipped_empty += 1
                continue

            code_text = normalize_newlines(str(code)).strip()
            if not code_text:
                skipped_empty += 1
                continue

            module_name = pick_module_name(str(vba_filename or ""), str(stream_path or ""))
            safe_name = sanitize_filename(module_name)

            base_path = out_dir / f"{OUT_PREFIX}{safe_name}.txt"
            out_path = unique_path(base_path, used_filenames)

            out_path.write_text(code_text + "\n", encoding="utf-8")
            written[module_name] = out_path
            exported_count += 1

            if not SUBFUNC_RE.search(code_text):
                no_subfunc_count += 1
                print(
                    f"Warning: exported module '{module_name}' but it contains no Sub/Function."
                )

    finally:
        # oletools versions differ; close() exists on newer versions
        close = getattr(vba, "close", None)
        if callable(close):
            close()

    print(f"Exported {exported_count} non-empty VBA modules to Mod_*.txt")
    print(f"Skipped {skipped_empty} empty code objects/modules.")
    if exported_count:
        print(f"Modules without Sub/Function: {no_subfunc_count}")


if __name__ == "__main__":
    main()

Du:
## TASK 3 – Implement Code to Extract Data from Excel via earlier produced .csv files

1. **Goal**: Create a simple script to create the extract of the relevant data (data_extract.py)
2. **Input for the script**: CSV files excelcell.csv, excelrange.csv (uploaded to LLM)  
3. **Output fo the script**  
   - var.csv  – Variables (sheet *Calculation*, A4:B9), contract‑specific  
   - tariff.csv – Tariff data (sheet *Calculation*, D4:E11), shared across multiple contracts  
   - limits.csv – Limits (sheet *Calculation*, G4:H5)  
   - tables.csv – MortalityTable (sheet *MortalityTables*, columns A–E, data starting in row 4, headers in row 3)  
   - tariff.py  – function **ModalSurcharge(PayFreq)** (Excel formula E12)  
4. **Procedure for the script**  
   - Read each listed cell block from the input files.  
   - Write CSVs exactly in the specified column formats.  
   - For CSVs, always include one column *Name* and one column *Value*.  
   - Implement ModalSurcharge(PayFreq) exactly according to the formula in cell E12.  
5. **Success Check**  
   - All files exist & have ≥ 1 data row (tables ≥ 100).  
   - import tariff; tariff.ModalSurcharge(12) returns the same value as Excel cell E12.

ChatGPT:

Nachgedacht für 48s
#!/usr/bin/env python3
"""
DATA_EXTRACT: Create normalized input datasets from excel exports.

Inputs (expected in project root):
  - excelcell.csv   columns: Sheet, Address, Formula, Value
  - excelrange.csv  columns: Sheet, Name, Address   (not required for the specific outputs, but validated)

If not found in CWD, the script also checks:
  - /mnt/data/excelcell.csv
  - /mnt/data/excelrange.csv

Outputs (written to project root / current working directory):
  - var.csv     (Name, Value) from Calculation!A4:B9
  - tariff.csv  (Name, Value) from Calculation!D4:E11
  - limits.csv  (Name, Value) from Calculation!G4:H5
  - tables.csv  (Name, Value) from MortalityTables!A:E (headers row 3, data from row 4)
  - tariff.py   implements ModalSurcharge(PayFreq) exactly per Calculation!E12

Success intent:
  - All files exist & have >= 1 data row (tables >= 100 expected for this workbook)
  - import tariff; tariff.ModalSurcharge(12) matches the Excel E12 logic.
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Dict, List, Tuple

# ----------------------------
# Paths
# ----------------------------
CANDIDATE_EXCELCELL = [Path("excelcell.csv"), Path("/mnt/data/excelcell.csv")]
CANDIDATE_EXCELRANGE = [Path("excelrange.csv"), Path("/mnt/data/excelrange.csv")]

OUT_VAR = Path("var.csv")
OUT_TARIFF = Path("tariff.csv")
OUT_LIMITS = Path("limits.csv")
OUT_TABLES = Path("tables.csv")
OUT_TARIFF_PY = Path("tariff.py")

A1_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.IGNORECASE)


def pick_existing(candidates: List[Path]) -> Path:
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(f"None of these input files exist: {', '.join(map(str, candidates))}")


# ----------------------------
# A1 helpers
# ----------------------------
def col_letters_to_num(letters: str) -> int:
    letters = letters.strip().upper()
    n = 0
    for ch in letters:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column letters: {letters}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def col_num_to_letters(n: int) -> str:
    if n < 1:
        raise ValueError(f"Invalid column number: {n}")
    out = []
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def a1(row: int, col: int) -> str:
    return f"{col_num_to_letters(col)}{row}"


# ----------------------------
# Load excelcell.csv into mapping
# ----------------------------
def load_excelcell(path: Path) -> Dict[Tuple[str, str], Dict[str, str]]:
    """
    Mapping:
      (sheet, address_upper) -> {"value": str, "formula": str}
    """
    out: Dict[Tuple[str, str], Dict[str, str]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        required = {"Sheet", "Address", "Formula", "Value"}
        if not r.fieldnames or not required.issubset(set(r.fieldnames)):
            raise ValueError(f"{path} must have columns {sorted(required)}; got {r.fieldnames}")

        for row in r:
            sheet = (row.get("Sheet") or "").strip()
            addr = (row.get("Address") or "").strip().upper()
            if not sheet or not addr:
                continue
            out[(sheet, addr)] = {
                "value": row.get("Value") or "",
                "formula": row.get("Formula") or "",
            }
    return out


def validate_excelrange(path: Path) -> None:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        required = {"Sheet", "Name", "Address"}
        if not r.fieldnames or not required.issubset(set(r.fieldnames)):
            raise ValueError(f"{path} must have columns {sorted(required)}; got {r.fieldnames}")


def get_value(cells: Dict[Tuple[str, str], Dict[str, str]], sheet: str, addr: str) -> str:
    rec = cells.get((sheet, addr.upper()))
    return "" if rec is None else (rec.get("value") or "")


def get_formula(cells: Dict[Tuple[str, str], Dict[str, str]], sheet: str, addr: str) -> str:
    rec = cells.get((sheet, addr.upper()))
    return "" if rec is None else (rec.get("formula") or "")


# ----------------------------
# Block extraction helpers
# ----------------------------
def write_name_value(path: Path, rows: List[Tuple[str, str]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Name", "Value"])
        for name, value in rows:
            name_s = (name or "").strip()
            value_s = (value or "").strip()
            if name_s == "" and value_s == "":
                continue
            w.writerow([name_s, value_s])


def extract_two_col_rows(
    cells: Dict[Tuple[str, str], Dict[str, str]],
    sheet: str,
    name_col: str,
    value_col: str,
    row_start: int,
    row_end: int,
) -> List[Tuple[str, str]]:
    nc = col_letters_to_num(name_col)
    vc = col_letters_to_num(value_col)
    out: List[Tuple[str, str]] = []
    for r in range(row_start, row_end + 1):
        name = get_value(cells, sheet, a1(r, nc))
        value = get_value(cells, sheet, a1(r, vc))
        if (name or "").strip() == "" and (value or "").strip() == "":
            continue
        out.append((name, value))
    return out


# ----------------------------
# MortalityTables extraction
# ----------------------------
def parse_row_col(addr: str) -> Tuple[int, int]:
    m = A1_RE.match(addr.strip().upper())
    if not m:
        raise ValueError(f"Bad A1: {addr}")
    col = col_letters_to_num(m.group(1))
    row = int(m.group(2))
    return row, col


def extract_mortality_tables(
    cells: Dict[Tuple[str, str], Dict[str, str]],
    sheet: str = "MortalityTables",
    header_row: int = 3,
    data_row_start: int = 4,
    col_start: str = "A",
    col_end: str = "E",
) -> List[Tuple[str, str]]:
    c1 = col_letters_to_num(col_start)
    c2 = col_letters_to_num(col_end)

    # headers from row 3
    headers: List[str] = []
    for c in range(c1, c2 + 1):
        h = (get_value(cells, sheet, a1(header_row, c)) or "").strip()
        headers.append(h if h else f"Col{col_num_to_letters(c)}")

    # find max data row that has any non-empty value within A:E
    max_row = 0
    for (sh, addr), rec in cells.items():
        if sh != sheet:
            continue
        try:
            r, c = parse_row_col(addr)
        except ValueError:
            continue
        if r < data_row_start or not (c1 <= c <= c2):
            continue
        v = (rec.get("value") or "").strip()
        f = (rec.get("formula") or "").strip()
        if v != "" or f != "":
            max_row = max(max_row, r)

    if max_row < data_row_start:
        return []

    out: List[Tuple[str, str]] = []
    for r in range(data_row_start, max_row + 1):
        row_dict: Dict[str, str] = {}
        any_nonempty = False
        first_val = ""
        for i, c in enumerate(range(c1, c2 + 1)):
            v = (get_value(cells, sheet, a1(r, c)) or "").strip()
            row_dict[headers[i]] = v
            if i == 0:
                first_val = v
            if v != "":
                any_nonempty = True

        if not any_nonempty:
            continue

        # Stable row key: include first column value if present, else row number
        key = first_val if first_val else f"ROW{r}"
        name = f"{key}|{r}"
        value = json.dumps(row_dict, ensure_ascii=False)
        out.append((name, value))

    return out


# ----------------------------
# ModalSurcharge extraction & codegen
# ----------------------------
E12_PATTERN = re.compile(
    r"""^=IF\s*\(\s*PayFreq\s*=\s*(\d+)\s*,\s*([\d.]+)%\s*,\s*IF\s*\(\s*PayFreq\s*=\s*(\d+)\s*,\s*([\d.]+)%\s*,\s*IF\s*\(\s*PayFreq\s*=\s*(\d+)\s*,\s*([\d.]+)%\s*,\s*0\s*\)\s*\)\s*\)\s*$""",
    re.IGNORECASE,
)


def render_tariff_py_from_e12(e12_formula: str) -> str:
    """
    Exact implementation for this workbook's E12 formula:
      =IF(PayFreq=2,2%,IF(PayFreq=4,3%,IF(PayFreq=12,5%,0)))
    We generate code directly from the formula, and validate it matches this pattern.
    """
    f = (e12_formula or "").strip()
    if not f:
        raise ValueError("Missing E12 formula text.")

    m = E12_PATTERN.match(f.replace(" ", ""))
    if not m:
        # Fallback: keep formula as comment and implement the known behavior only if it matches expected set.
        # But we fail hard to avoid silent mismatches.
        raise ValueError(f"Unexpected E12 formula format: {f!r}")

    pf1, p1, pf2, p2, pf3, p3 = m.groups()
    mapping = {
        int(pf1): float(p1) / 100.0,
        int(pf2): float(p2) / 100.0,
        int(pf3): float(p3) / 100.0,
    }

    return f'''"""
Auto-generated by data_extract.py from excelcell.csv.

Implements ModalSurcharge(PayFreq) exactly as the Excel formula in Calculation!E12:
  {f}
"""

from __future__ import annotations


EXCEL_E12_FORMULA = {f!r}
_MODAL_MAP = {mapping!r}


def ModalSurcharge(PayFreq: int) -> float:
    """
    Modal surcharge for payment frequency.

    Mirrors the nested IF in Excel cell Calculation!E12.
    """
    try:
        pf = int(PayFreq)
    except Exception as e:
        raise TypeError("PayFreq must be convertible to int") from e
    return float(_MODAL_MAP.get(pf, 0.0))
'''


# ----------------------------
# Main
# ----------------------------
def main() -> None:
    excelcell = pick_existing(CANDIDATE_EXCELCELL)
    excelrange = pick_existing(CANDIDATE_EXCELRANGE)

    validate_excelrange(excelrange)
    cells = load_excelcell(excelcell)

    # var.csv: Calculation A4:B9
    var_rows = extract_two_col_rows(cells, "Calculation", "A", "B", 4, 9)
    write_name_value(OUT_VAR, var_rows)

    # tariff.csv: Calculation D4:E11
    tariff_rows = extract_two_col_rows(cells, "Calculation", "D", "E", 4, 11)
    write_name_value(OUT_TARIFF, tariff_rows)

    # limits.csv: Calculation G4:H5
    limits_rows = extract_two_col_rows(cells, "Calculation", "G", "H", 4, 5)
    write_name_value(OUT_LIMITS, limits_rows)

    # tables.csv: MortalityTables A:E, headers row 3, data from row 4
    tables_rows = extract_mortality_tables(cells)
    write_name_value(OUT_TABLES, tables_rows)

    # tariff.py: ModalSurcharge(PayFreq) from Calculation!E12 formula
    e12_formula = get_formula(cells, "Calculation", "E12")
    if not e12_formula.strip():
        raise ValueError("Could not find Calculation!E12 formula in excelcell.csv (missing or empty Formula column).")
    OUT_TARIFF_PY.write_text(render_tariff_py_from_e12(e12_formula), encoding="utf-8")

    # Minimal runtime checks (non-fatal prints)
    def data_rows(p: Path) -> int:
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            return max(0, sum(1 for _ in f) - 1)

    print(f"Inputs: {excelcell} , {excelrange}")
    print(f"Wrote {OUT_VAR} rows={data_rows(OUT_VAR)}")
    print(f"Wrote {OUT_TARIFF} rows={data_rows(OUT_TARIFF)}")
    print(f"Wrote {OUT_LIMITS} rows={data_rows(OUT_LIMITS)}")
    print(f"Wrote {OUT_TABLES} rows={data_rows(OUT_TABLES)}")
    print(f"Wrote {OUT_TARIFF_PY} (ModalSurcharge from Calculation!E12)")


if __name__ == "__main__":
    main()

Du:
## TASK 4 – Implement Code to Generate Test Fixtures

1. **Goal**: Creat tests (tests/conftest.py, tests/test_data_roundtrip.py)
2. **Input for LLM**: Existing CSV files (var.csv, tariff.csv, tariff.py, limits.csv, tables.csv)  
3. **Input for the script**: Existing CSV files (var.csv, tariff.csv, tariff.py, limits.csv, tables.csv)  
4. **Output**: Folder tests/ with pytest fixtures  
5. **Procedure**  
   - LLM create two code blocks
   - conftest.py sets up temp directory & mini CSV samples.  
   - Create smoke test test_data_roundtrip.py that reads each CSV and counts columns.  
6. **Success Check**  
   - pytest -q runs green (0 errors, 0 failures).

ChatGPT:
# tests/conftest.py
from __future__ import annotations

import csv
import shutil
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FALLBACK_DIR = Path("/mnt/data")


def _source_file(name: str) -> Path:
    """Prefer project root, fall back to /mnt/data (used by some LLM environments)."""
    p = PROJECT_ROOT / name
    if p.exists():
        return p
    p2 = FALLBACK_DIR / name
    if p2.exists():
        return p2
    raise FileNotFoundError(f"Required input file not found in project root or /mnt/data: {name}")


def _mini_csv(src: Path, dst: Path, max_rows: int) -> None:
    """Copy CSV with header + up to max_rows data rows."""
    with src.open("r", encoding="utf-8-sig", newline="") as fin:
        reader = csv.reader(fin)
        try:
            header = next(reader)
        except StopIteration as e:
            raise ValueError(f"CSV has no header: {src}") from e

        rows = [header]
        for i, row in enumerate(reader):
            if i >= max_rows:
                break
            rows.append(row)

    with dst.open("w", encoding="utf-8", newline="") as fout:
        writer = csv.writer(fout)
        writer.writerows(rows)


@pytest.fixture()
def data_dir(tmp_path: Path) -> Path:
    """
    Temp dir containing small, deterministic samples of the product data.
    """
    mini_specs = {
        "var.csv": 50,
        "tariff.csv": 50,
        "limits.csv": 50,
        "tables.csv": 200,
    }

    for filename, max_rows in mini_specs.items():
        src = _source_file(filename)
        _mini_csv(src, tmp_path / filename, max_rows=max_rows)

    shutil.copyfile(_source_file("tariff.py"), tmp_path / "tariff.py")
    return tmp_path
# tests/test_data_roundtrip.py
from __future__ import annotations

import csv
import importlib.util
from pathlib import Path


def _read_header(path: Path) -> list[str]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            return next(r)
        except StopIteration:
            return []


def _count_rows(path: Path) -> int:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            next(r)  # header
        except StopIteration:
            return 0
        return sum(1 for _ in r)


def test_csv_roundtrip_smoke(data_dir: Path) -> None:
    for filename in ("var.csv", "tariff.csv", "limits.csv", "tables.csv"):
        p = data_dir / filename
        assert p.exists(), f"Missing file: {p}"

        header = _read_header(p)
        assert header, f"{filename}: missing header row"
        assert "Name" in header and "Value" in header, f"{filename}: expected Name/Value columns, got {header}"
        assert len(header) >= 2, f"{filename}: expected >=2 columns, got {len(header)}"

        assert _count_rows(p) >= 1, f"{filename}: expected at least 1 data row"

    # tables.csv should be non-trivial even in mini sample
    assert _count_rows(data_dir / "tables.csv") >= 10


def test_tariff_module_import_and_modal_surcharge(data_dir: Path) -> None:
    tariff_path = data_dir / "tariff.py"
    assert tariff_path.exists()

    spec = importlib.util.spec_from_file_location("tariff_testcopy", str(tariff_path))
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]

    assert hasattr(mod, "ModalSurcharge")
    v = mod.ModalSurcharge(12)
    assert isinstance(v, float)

Du:
## TASK 5A – Implement Code to Translate Base Functions

1. **Goal**: Translate VBA code into python
2. **Input for LLM**: All Mod_*.txt from TASK 2  
3. **Output of LLM**: basfunct.py  
4. **Procedure**  
   - Each VBA Function/Procedure is mapped 1‑to‑1 to a Python def.  
   - Use pandas for table/CSV access.  
   - No function body may end with pass.  
   - Available data sources for the python script: excelcell.csv, excelrange.csv, var.csv, tariff.csv, limits.csv, tables.csv.


ChatGPT:
# basfunct.py
"""
BASISFUNCT – 1-to-1 port of VBA base functions to Python.

Source VBA modules:
- Mod_mPresentValues.txt
- Mod_mCommValues.txt
- Mod_mConstants.txt

Data sources (CSV):
- tables.csv (MortalityTables) is used for Act_qx / commutation functions.

Notes:
- Uses pandas for CSV access.
- Caching mirrors VBA's Scripting.Dictionary usage.
- Excel/VBA rounding differs from Python's bankers rounding; we implement Excel-like ROUND.
- Fix: avoids deprecated pd.to_numeric(errors="ignore") to silence FutureWarning.
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd

# ----------------------------
# Constants (from mConstants)
# ----------------------------
round_lx: int = 16
round_tx: int = 16
round_Dx: int = 16
round_Cx: int = 16
round_Nx: int = 16
round_Mx: int = 16
round_Rx: int = 16
max_Age: int = 123


# ----------------------------
# Data access
# ----------------------------
@dataclass
class DataRepo:
    root: Path = Path.cwd()
    _cache: Dict[str, pd.DataFrame] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self._cache is None:
            self._cache = {}

    def _path(self, name: str) -> Path:
        p = self.root / name
        if p.exists():
            return p
        p2 = Path("/mnt/data") / name
        if p2.exists():
            return p2
        raise FileNotFoundError(f"Missing data source: {name}")

    def read_csv(self, name: str) -> pd.DataFrame:
        if name not in self._cache:
            self._cache[name] = pd.read_csv(self._path(name), encoding="utf-8")
        return self._cache[name]

    @staticmethod
    def _safe_numeric_coerce(series: pd.Series) -> pd.Series:
        """
        Convert string-like series to numeric where possible.
        - Empty strings / 'nan' / 'None' become NaN.
        - Values that cannot be converted remain as their original (string) value.
        This avoids deprecated `errors="ignore"` behavior while preserving non-numeric text.
        """
        if series.dtype != object:
            return series

        s_str = series.astype(str)
        empties = s_str.str.strip().isin(["", "nan", "None"])
        s_clean = s_str.where(~empties, None)

        numeric = pd.to_numeric(s_clean, errors="coerce")
        # Keep original where conversion failed but original wasn't empty
        failed = numeric.isna() & ~empties

        # If nothing failed, return numeric
        if not failed.any():
            return numeric

        # Mixed: preserve originals for failed values, numeric for successful values
        out = series.copy()
        out.loc[~failed] = numeric.loc[~failed]
        return out

    def mortality_tables_df(self) -> pd.DataFrame:
        """
        Build a DataFrame from tables.csv (Name, Value where Value is JSON per row).
        Expected JSON keys correspond to MortalityTables headers.
        """
        df = self.read_csv("tables.csv")
        if "Value" not in df.columns:
            raise ValueError("tables.csv must have column 'Value'")

        rows = []
        for v in df["Value"].astype(str):
            try:
                rows.append(json.loads(v))
            except json.JSONDecodeError:
                # allow already-plain rows (shouldn't happen)
                rows.append({"_raw": v})

        tdf = pd.DataFrame(rows)

        # Coerce numeric columns safely (no FutureWarning)
        for c in tdf.columns:
            tdf[c] = self._safe_numeric_coerce(tdf[c])

        return tdf


_DATA = DataRepo()


# ----------------------------
# Excel-like rounding helpers
# ----------------------------
def _excel_round(x: float, digits: int = 0) -> float:
    """
    Excel/VBA WorksheetFunction.Round: halves away from zero.
    Python round() is bankers rounding; do NOT use it here.
    """
    if digits >= 0:
        factor = 10.0**digits
        return math.copysign(math.floor(abs(x) * factor + 0.5) / factor, x)
    factor = 10.0 ** (-digits)
    return math.copysign(math.floor(abs(x) / factor + 0.5) * factor, x)


# ----------------------------
# Cache (from mCommValues)
# ----------------------------
cache: Optional[Dict[str, float]] = None


def InitializeCache() -> None:
    """Create a new Dictionary object (Python dict)."""
    global cache
    cache = {}


# ----------------------------
# Mortality / commutation functions (from mCommValues)
# ----------------------------
def Act_qx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    """
    Equivalent to VBA:
      Select Case TableId: "DAV1994_T", "DAV2008_T"
      tableVector = UCase(TableId) & "_" & Sex
      Index(m_Tables, Age+1, Match(tableVector, v_Tables, 0))
    """
    sex = (Sex or "").upper()
    if sex != "M":
        sex = "F"

    table_id = (TableId or "").upper()
    if table_id not in ("DAV1994_T", "DAV2008_T"):
        raise ValueError(f"Unsupported TableId: {TableId}")

    table_vector = f"{table_id}_{sex}"

    tdf = _DATA.mortality_tables_df()

    # Find an "Age" column (common), otherwise assume the first column is age-like.
    age_col = None
    for cand in ("Age", "AGE", "alter", "ALTER"):
        if cand in tdf.columns:
            age_col = cand
            break

    if table_vector not in tdf.columns:
        raise KeyError(f"Column '{table_vector}' not found in tables.csv-derived DataFrame")

    if age_col is not None:
        match = tdf.loc[tdf[age_col] == Age, table_vector]
        if match.empty:
            # fallback: treat row index as age (0-based)
            if 0 <= Age < len(tdf):
                return float(tdf.iloc[Age][table_vector])
            raise IndexError(f"Age {Age} not found in Age column and out of bounds.")
        return float(match.iloc[0])

    # No age column: assume 0-based row corresponds to Age=0
    if 0 <= Age < len(tdf):
        return float(tdf.iloc[Age][table_vector])
    raise IndexError(f"Age {Age} out of bounds for tables data (len={len(tdf)}).")


def Vec_lx(
    EndAge: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of lx; if EndAge = -1 then it is created up to max_Age."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    vec[0] = 1_000_000.0
    for i in range(1, limit + 1):
        vec[i] = vec[i - 1] * (1.0 - Act_qx(i - 1, Sex, TableId, BirthYear, RetirementAge, Layer))
        vec[i] = float(_excel_round(vec[i], round_lx))
    return vec


def Act_lx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    vec = Vec_lx(Age, Sex, TableId, BirthYear, RetirementAge, Layer)
    return float(vec[Age])


def Vec_tx(
    EndAge: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of tx (# deaths)."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    temp_lx = Vec_lx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit):
        vec[i] = temp_lx[i] - temp_lx[i + 1]
        vec[i] = float(_excel_round(vec[i], round_tx))
    return vec


def Act_tx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    vec = Vec_tx(Age, Sex, TableId, BirthYear, RetirementAge, Layer)
    return float(vec[Age])


def Vec_Dx(
    EndAge: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Dx."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    v = 1.0 / (1.0 + float(InterestRate))
    temp_lx = Vec_lx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit + 1):
        vec[i] = temp_lx[i] * (v**i)
        vec[i] = float(_excel_round(vec[i], round_Dx))
    return vec


def BuildCacheKey(
    Kind: str,
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int,
    RetirementAge: int,
    Layer: int,
) -> str:
    return f"{Kind}_{Age}_{Sex}_{TableId}_{InterestRate}_{BirthYear}_{RetirementAge}_{Layer}"


def Act_Dx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Dx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Cx(
    EndAge: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Cx."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    v = 1.0 / (1.0 + float(InterestRate))
    temp_tx = Vec_tx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit):
        vec[i] = temp_tx[i] * (v ** (i + 1))
        vec[i] = float(_excel_round(vec[i], round_Cx))
    return vec


def Act_Cx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Cx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Cx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Nx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Nx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_dx = Vec_Dx(-1, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_dx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_dx[i]
        vec[i] = float(_excel_round(vec[i], round_Dx))  # kept as in original
    return vec


def Act_Nx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Nx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Nx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Mx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Mx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_cx = Vec_Cx(-1, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_cx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_cx[i]
        vec[i] = float(_excel_round(vec[i], round_Mx))
    return vec


def Act_Mx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Mx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Mx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Rx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Rx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_mx = Vec_Mx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_mx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_mx[i]
        vec[i] = float(_excel_round(vec[i], round_Rx))
    return vec


def Act_Rx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Rx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Rx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Act_AgeCalculation(BirthDate: date, ValuationDate: date, Method: str) -> int:
    """Age calculation based on calendar-year method (K) or half-year method (H)."""
    method = Method if Method == "K" else "H"

    # Accept datetime as well
    if isinstance(BirthDate, datetime):
        BirthDate = BirthDate.date()
    if isinstance(ValuationDate, datetime):
        ValuationDate = ValuationDate.date()

    y_birth = BirthDate.year
    y_val = ValuationDate.year
    m_birth = BirthDate.month
    m_val = ValuationDate.month

    if method == "K":
        return int(y_val - y_birth)
    # "H"
    return int(math.floor(y_val - y_birth + (1.0 / 12.0) * (m_val - m_birth + 5)))


# ----------------------------
# Present value functions (from mPresentValues)
# ----------------------------
def Act_DeductionTerm(k: int, InterestRate: float) -> float:
    """Deduction term."""
    acc = 0.0
    if k > 0:
        for l in range(0, k):
            acc += (l / k) / (1.0 + (l / k) * float(InterestRate))
        acc = acc * (1.0 + float(InterestRate)) / k
    return float(acc)


def Act_ax_k(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k > 0:
        return float(
            Act_Nx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            - Act_DeductionTerm(k, float(InterestRate))
        )
    return 0.0


def Act_axn_k(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k <= 0:
        return 0.0

    nx_age = Act_Nx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    nx_agen = Act_Nx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    dx_age = Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    dx_agen = Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)

    return float(
        (nx_age - nx_agen) / dx_age
        - Act_DeductionTerm(k, float(InterestRate)) * (1.0 - dx_agen / dx_age)
    )


def Act_nax_k(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k > 0:
        return float(
            Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            * Act_ax_k(Age + n, Sex, TableId, float(InterestRate), k, BirthYear, RetirementAge, Layer)
        )
    return 0.0


def act_nGrAx(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    return float(
        (
            Act_Mx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            - Act_Mx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
        )
        / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    )


def act_nGrEx(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    return float(
        Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
        / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    )


def Act_ag_k(g: int, InterestRate: float, k: int) -> float:
    v = 1.0 / (1.0 + float(InterestRate))
    if k > 0:
        if float(InterestRate) > 0:
            return float((1.0 - v**g) / (1.0 - v) - Act_DeductionTerm(k, float(InterestRate)) * (1.0 - v**g))
        return float(g)
    return 0.0
Du:
## TASK 5B – Implement Code for Function Parity Test

**Success Check**: LLM creates `tests/test_func_parity.py` that  
- collects all **public** VBA names (`Function`/`Sub` without `Private`),  
- scans Python `def` names in `basfunct.py` (helpers may be ignored),  
- and verifies that for each VBA name exactly one Python function exists.  
Passing = `pytest -q` runs completely green.

ChatGPT:
# tests/test_func_parity.py
from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Iterable, Set

# VBA: Public if not explicitly Private
VBA_DECL_RE = re.compile(
    r"(?im)^\s*(?:(public|private|friend)\s+)?(function|sub)\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)

# Some VBA members we intentionally ignore (e.g., worksheet event stubs, if any appear)
VBA_IGNORE_NAMES: Set[str] = {
    # add known non-base / event handlers here if they appear in Mod_*.txt
}

# Python helpers inside basfunct.py we don't want to treat as "ported VBA names"
PY_IGNORE_NAMES: Set[str] = {
    "DataRepo",
    "_excel_round",
}


def _find_module_txt_files() -> list[Path]:
    # Prefer project root, but support /mnt/data (LLM sandbox)
    roots = [Path.cwd(), Path("/mnt/data")]
    files: list[Path] = []
    for root in roots:
        files.extend(sorted(root.glob("Mod_*.txt")))
    # de-dupe by resolved path string
    uniq: dict[str, Path] = {}
    for p in files:
        try:
            uniq[str(p.resolve())] = p
        except Exception:
            uniq[str(p)] = p
    return list(uniq.values())


def _vba_public_names_from_text(text: str) -> Set[str]:
    names: Set[str] = set()
    for m in VBA_DECL_RE.finditer(text):
        vis = (m.group(1) or "").strip().lower()
        kind = (m.group(2) or "").strip().lower()
        name = (m.group(3) or "").strip()

        # Only Functions/Subs
        if kind not in ("function", "sub"):
            continue

        # Exclude Private
        if vis == "private":
            continue

        if name in VBA_IGNORE_NAMES:
            continue

        names.add(name)
    return names


def _collect_vba_public_names() -> Set[str]:
    mod_files = _find_module_txt_files()
    assert mod_files, "No Mod_*.txt files found (run TASK 2 VBA export)."

    names: Set[str] = set()
    for p in mod_files:
        text = p.read_text(encoding="utf-8", errors="ignore")
        names |= _vba_public_names_from_text(text)
    assert names, "No public VBA Function/Sub names found in Mod_*.txt files."
    return names


def _collect_python_def_names(basfunct_path: Path) -> Set[str]:
    src = basfunct_path.read_text(encoding="utf-8")
    tree = ast.parse(src, filename=str(basfunct_path))
    names: Set[str] = set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            if node.name in PY_IGNORE_NAMES:
                continue
            if node.name.startswith("_"):
                continue
            names.add(node.name)
    return names


def test_public_vba_names_have_python_defs() -> None:
    vba_names = _collect_vba_public_names()

    # Locate basfunct.py in project root (or /mnt/data fallback)
    basfunct_candidates = [Path.cwd() / "basfunct.py", Path("/mnt/data") / "basfunct.py"]
    basfunct_path = next((p for p in basfunct_candidates if p.exists()), None)
    assert basfunct_path is not None, "basfunct.py not found (create it in TASK 5A)."

    py_names = _collect_python_def_names(basfunct_path)

    missing = sorted(n for n in vba_names if n not in py_names)
    assert not missing, f"Missing Python defs for VBA public names: {missing}"


def test_no_duplicate_python_defs_for_vba_names() -> None:
    """
    Ensures each VBA public name maps to exactly one Python def.
    (In Python, duplicate def names in the same module would overwrite; we detect that by AST only
    yielding final name set. So we do a stricter check by scanning raw text for 'def <name>(' counts.)
    """
    vba_names = _collect_vba_public_names()

    basfunct_candidates = [Path.cwd() / "basfunct.py", Path("/mnt/data") / "basfunct.py"]
    basfunct_path = next((p for p in basfunct_candidates if p.exists()), None)
    assert basfunct_path is not None, "basfunct.py not found (create it in TASK 5A)."

    src = basfunct_path.read_text(encoding="utf-8")

    duplicates = []
    for name in sorted(vba_names):
        # exact 'def Name(' occurrences
        cnt = len(re.findall(rf"(?m)^\s*def\s+{re.escape(name)}\s*\(", src))
        if cnt != 1:
            duplicates.append((name, cnt))

    assert not duplicates, f"Expected exactly one Python def for each VBA name; mismatches: {duplicates}"


---

## 2. Already Implemented Python Artifacts - Extraction Tools

### FILENAME: data_extract.py
CONTENT:

/not needed/


---


### FILENAME: excel_to_text.py
CONTENT:

/not needed/

---


### FILENAME: vba_to_text.py
CONTENT:

/not needed/

---

## 3. Already Implemented Python Artifacts - Main Code

### FILENAME: basfunct.py
CONTENT:

# basfunct.py
"""
BASISFUNCT – 1-to-1 port of VBA base functions to Python.

Source VBA modules:
- Mod_mPresentValues.txt
- Mod_mCommValues.txt
- Mod_mConstants.txt

Data sources (CSV):
- tables.csv (MortalityTables) is used for Act_qx / commutation functions.

Notes:
- Uses pandas for CSV access.
- Caching mirrors VBA's Scripting.Dictionary usage.
- Excel/VBA rounding differs from Python's bankers rounding; we implement Excel-like ROUND.
- Fix: avoids deprecated pd.to_numeric(errors="ignore") to silence FutureWarning.
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd

# ----------------------------
# Constants (from mConstants)
# ----------------------------
round_lx: int = 16
round_tx: int = 16
round_Dx: int = 16
round_Cx: int = 16
round_Nx: int = 16
round_Mx: int = 16
round_Rx: int = 16
max_Age: int = 123


# ----------------------------
# Data access
# ----------------------------
@dataclass
class DataRepo:
    root: Path = Path.cwd()
    _cache: Dict[str, pd.DataFrame] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self._cache is None:
            self._cache = {}

    def _path(self, name: str) -> Path:
        p = self.root / name
        if p.exists():
            return p
        p2 = Path("/mnt/data") / name
        if p2.exists():
            return p2
        raise FileNotFoundError(f"Missing data source: {name}")

    def read_csv(self, name: str) -> pd.DataFrame:
        if name not in self._cache:
            self._cache[name] = pd.read_csv(self._path(name), encoding="utf-8")
        return self._cache[name]

    @staticmethod
    def _safe_numeric_coerce(series: pd.Series) -> pd.Series:
        """
        Convert string-like series to numeric where possible.
        - Empty strings / 'nan' / 'None' become NaN.
        - Values that cannot be converted remain as their original (string) value.
        This avoids deprecated `errors="ignore"` behavior while preserving non-numeric text.
        """
        if series.dtype != object:
            return series

        s_str = series.astype(str)
        empties = s_str.str.strip().isin(["", "nan", "None"])
        s_clean = s_str.where(~empties, None)

        numeric = pd.to_numeric(s_clean, errors="coerce")
        # Keep original where conversion failed but original wasn't empty
        failed = numeric.isna() & ~empties

        # If nothing failed, return numeric
        if not failed.any():
            return numeric

        # Mixed: preserve originals for failed values, numeric for successful values
        out = series.copy()
        out.loc[~failed] = numeric.loc[~failed]
        return out

    def mortality_tables_df(self) -> pd.DataFrame:
        """
        Build a DataFrame from tables.csv (Name, Value where Value is JSON per row).
        Expected JSON keys correspond to MortalityTables headers.
        """
        df = self.read_csv("tables.csv")
        if "Value" not in df.columns:
            raise ValueError("tables.csv must have column 'Value'")

        rows = []
        for v in df["Value"].astype(str):
            try:
                rows.append(json.loads(v))
            except json.JSONDecodeError:
                # allow already-plain rows (shouldn't happen)
                rows.append({"_raw": v})

        tdf = pd.DataFrame(rows)

        # Coerce numeric columns safely (no FutureWarning)
        for c in tdf.columns:
            tdf[c] = self._safe_numeric_coerce(tdf[c])

        return tdf


_DATA = DataRepo()


# ----------------------------
# Excel-like rounding helpers
# ----------------------------
def _excel_round(x: float, digits: int = 0) -> float:
    """
    Excel/VBA WorksheetFunction.Round: halves away from zero.
    Python round() is bankers rounding; do NOT use it here.
    """
    if digits >= 0:
        factor = 10.0**digits
        return math.copysign(math.floor(abs(x) * factor + 0.5) / factor, x)
    factor = 10.0 ** (-digits)
    return math.copysign(math.floor(abs(x) / factor + 0.5) * factor, x)


# ----------------------------
# Cache (from mCommValues)
# ----------------------------
cache: Optional[Dict[str, float]] = None


def InitializeCache() -> None:
    """Create a new Dictionary object (Python dict)."""
    global cache
    cache = {}


# ----------------------------
# Mortality / commutation functions (from mCommValues)
# ----------------------------
def Act_qx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    """
    Equivalent to VBA:
      Select Case TableId: "DAV1994_T", "DAV2008_T"
      tableVector = UCase(TableId) & "_" & Sex
      Index(m_Tables, Age+1, Match(tableVector, v_Tables, 0))
    """
    sex = (Sex or "").upper()
    if sex != "M":
        sex = "F"

    table_id = (TableId or "").upper()
    if table_id not in ("DAV1994_T", "DAV2008_T"):
        raise ValueError(f"Unsupported TableId: {TableId}")

    table_vector = f"{table_id}_{sex}"

    tdf = _DATA.mortality_tables_df()

    # Find an "Age" column (common), otherwise assume the first column is age-like.
    age_col = None
    for cand in ("Age", "AGE", "alter", "ALTER"):
        if cand in tdf.columns:
            age_col = cand
            break

    if table_vector not in tdf.columns:
        raise KeyError(f"Column '{table_vector}' not found in tables.csv-derived DataFrame")

    if age_col is not None:
        match = tdf.loc[tdf[age_col] == Age, table_vector]
        if match.empty:
            # fallback: treat row index as age (0-based)
            if 0 <= Age < len(tdf):
                return float(tdf.iloc[Age][table_vector])
            raise IndexError(f"Age {Age} not found in Age column and out of bounds.")
        return float(match.iloc[0])

    # No age column: assume 0-based row corresponds to Age=0
    if 0 <= Age < len(tdf):
        return float(tdf.iloc[Age][table_vector])
    raise IndexError(f"Age {Age} out of bounds for tables data (len={len(tdf)}).")


def Vec_lx(
    EndAge: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of lx; if EndAge = -1 then it is created up to max_Age."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    vec[0] = 1_000_000.0
    for i in range(1, limit + 1):
        vec[i] = vec[i - 1] * (1.0 - Act_qx(i - 1, Sex, TableId, BirthYear, RetirementAge, Layer))
        vec[i] = float(_excel_round(vec[i], round_lx))
    return vec


def Act_lx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    vec = Vec_lx(Age, Sex, TableId, BirthYear, RetirementAge, Layer)
    return float(vec[Age])


def Vec_tx(
    EndAge: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of tx (# deaths)."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    temp_lx = Vec_lx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit):
        vec[i] = temp_lx[i] - temp_lx[i + 1]
        vec[i] = float(_excel_round(vec[i], round_tx))
    return vec


def Act_tx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    vec = Vec_tx(Age, Sex, TableId, BirthYear, RetirementAge, Layer)
    return float(vec[Age])


def Vec_Dx(
    EndAge: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Dx."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    v = 1.0 / (1.0 + float(InterestRate))
    temp_lx = Vec_lx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit + 1):
        vec[i] = temp_lx[i] * (v**i)
        vec[i] = float(_excel_round(vec[i], round_Dx))
    return vec


def BuildCacheKey(
    Kind: str,
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int,
    RetirementAge: int,
    Layer: int,
) -> str:
    return f"{Kind}_{Age}_{Sex}_{TableId}_{InterestRate}_{BirthYear}_{RetirementAge}_{Layer}"


def Act_Dx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Dx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Cx(
    EndAge: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Cx."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    v = 1.0 / (1.0 + float(InterestRate))
    temp_tx = Vec_tx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit):
        vec[i] = temp_tx[i] * (v ** (i + 1))
        vec[i] = float(_excel_round(vec[i], round_Cx))
    return vec


def Act_Cx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Cx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Cx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Nx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Nx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_dx = Vec_Dx(-1, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_dx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_dx[i]
        vec[i] = float(_excel_round(vec[i], round_Dx))  # kept as in original
    return vec


def Act_Nx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Nx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Nx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Mx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Mx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_cx = Vec_Cx(-1, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_cx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_cx[i]
        vec[i] = float(_excel_round(vec[i], round_Mx))
    return vec


def Act_Mx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Mx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Mx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Rx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Rx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_mx = Vec_Mx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_mx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_mx[i]
        vec[i] = float(_excel_round(vec[i], round_Rx))
    return vec


def Act_Rx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Rx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Rx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Act_AgeCalculation(BirthDate: date, ValuationDate: date, Method: str) -> int:
    """Age calculation based on calendar-year method (K) or half-year method (H)."""
    method = Method if Method == "K" else "H"

    # Accept datetime as well
    if isinstance(BirthDate, datetime):
        BirthDate = BirthDate.date()
    if isinstance(ValuationDate, datetime):
        ValuationDate = ValuationDate.date()

    y_birth = BirthDate.year
    y_val = ValuationDate.year
    m_birth = BirthDate.month
    m_val = ValuationDate.month

    if method == "K":
        return int(y_val - y_birth)
    # "H"
    return int(math.floor(y_val - y_birth + (1.0 / 12.0) * (m_val - m_birth + 5)))


# ----------------------------
# Present value functions (from mPresentValues)
# ----------------------------
def Act_DeductionTerm(k: int, InterestRate: float) -> float:
    """Deduction term."""
    acc = 0.0
    if k > 0:
        for l in range(0, k):
            acc += (l / k) / (1.0 + (l / k) * float(InterestRate))
        acc = acc * (1.0 + float(InterestRate)) / k
    return float(acc)


def Act_ax_k(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k > 0:
        return float(
            Act_Nx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            - Act_DeductionTerm(k, float(InterestRate))
        )
    return 0.0


def Act_axn_k(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k <= 0:
        return 0.0

    nx_age = Act_Nx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    nx_agen = Act_Nx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    dx_age = Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    dx_agen = Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)

    return float(
        (nx_age - nx_agen) / dx_age
        - Act_DeductionTerm(k, float(InterestRate)) * (1.0 - dx_agen / dx_age)
    )


def Act_nax_k(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k > 0:
        return float(
            Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            * Act_ax_k(Age + n, Sex, TableId, float(InterestRate), k, BirthYear, RetirementAge, Layer)
        )
    return 0.0


def act_nGrAx(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    return float(
        (
            Act_Mx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            - Act_Mx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
        )
        / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    )


def act_nGrEx(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    return float(
        Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
        / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    )


def Act_ag_k(g: int, InterestRate: float, k: int) -> float:
    v = 1.0 / (1.0 + float(InterestRate))
    if k > 0:
        if float(InterestRate) > 0:
            return float((1.0 - v**g) / (1.0 - v) - Act_DeductionTerm(k, float(InterestRate)) * (1.0 - v**g))
        return float(g)
    return 0.0


---

### FILENAME: tariff.py
CONTENT:

"""
Auto-generated by data_extract.py from excelcell.csv.

Implements ModalSurcharge(PayFreq) exactly as the Excel formula in Calculation!E12:
  =IF(PayFreq=2,2%,IF(PayFreq=4,3%,IF(PayFreq=12,5%,0)))
"""

from __future__ import annotations


EXCEL_E12_FORMULA = '=IF(PayFreq=2,2%,IF(PayFreq=4,3%,IF(PayFreq=12,5%,0)))'
_MODAL_MAP = {2: 0.02, 4: 0.03, 12: 0.05}


def ModalSurcharge(PayFreq: int) -> float:
    """
    Modal surcharge for payment frequency.

    Mirrors the nested IF in Excel cell Calculation!E12.
    """
    try:
        pf = int(PayFreq)
    except Exception as e:
        raise TypeError("PayFreq must be convertible to int") from e
    return float(_MODAL_MAP.get(pf, 0.0))


---

### FILENAME: outfunc.py
CONTENT:

/not created yet - part of the excerce/

---

### FILENAME: run_calc.py
CONTENT:

/not created yet - part of the excerce/

---


## 4. Already Implemented Python Artifacts - Tests

### FILENAME: conftest.py
CONTENT:

# tests/conftest.py
from __future__ import annotations

import csv
import shutil
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FALLBACK_DIR = Path("/mnt/data")


def _source_file(name: str) -> Path:
    """Prefer project root, fall back to /mnt/data (used by some LLM environments)."""
    p = PROJECT_ROOT / name
    if p.exists():
        return p
    p2 = FALLBACK_DIR / name
    if p2.exists():
        return p2
    raise FileNotFoundError(f"Required input file not found in project root or /mnt/data: {name}")


def _mini_csv(src: Path, dst: Path, max_rows: int) -> None:
    """Copy CSV with header + up to max_rows data rows."""
    with src.open("r", encoding="utf-8-sig", newline="") as fin:
        reader = csv.reader(fin)
        try:
            header = next(reader)
        except StopIteration as e:
            raise ValueError(f"CSV has no header: {src}") from e

        rows = [header]
        for i, row in enumerate(reader):
            if i >= max_rows:
                break
            rows.append(row)

    with dst.open("w", encoding="utf-8", newline="") as fout:
        writer = csv.writer(fout)
        writer.writerows(rows)


@pytest.fixture()
def data_dir(tmp_path: Path) -> Path:
    """
    Temp dir containing small, deterministic samples of the product data.
    """
    mini_specs = {
        "var.csv": 50,
        "tariff.csv": 50,
        "limits.csv": 50,
        "tables.csv": 200,
    }

    for filename, max_rows in mini_specs.items():
        src = _source_file(filename)
        _mini_csv(src, tmp_path / filename, max_rows=max_rows)

    shutil.copyfile(_source_file("tariff.py"), tmp_path / "tariff.py")
    return tmp_path


---

### FILENAME: test_data_roundtrip.py
CONTENT:

# tests/test_data_roundtrip.py
from __future__ import annotations

import csv
import importlib.util
from pathlib import Path


def _read_header(path: Path) -> list[str]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            return next(r)
        except StopIteration:
            return []


def _count_rows(path: Path) -> int:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            next(r)  # header
        except StopIteration:
            return 0
        return sum(1 for _ in r)


def test_csv_roundtrip_smoke(data_dir: Path) -> None:
    for filename in ("var.csv", "tariff.csv", "limits.csv", "tables.csv"):
        p = data_dir / filename
        assert p.exists(), f"Missing file: {p}"

        header = _read_header(p)
        assert header, f"{filename}: missing header row"
        assert "Name" in header and "Value" in header, f"{filename}: expected Name/Value columns, got {header}"
        assert len(header) >= 2, f"{filename}: expected >=2 columns, got {len(header)}"

        assert _count_rows(p) >= 1, f"{filename}: expected at least 1 data row"

    # tables.csv should be non-trivial even in mini sample
    assert _count_rows(data_dir / "tables.csv") >= 10


def test_tariff_module_import_and_modal_surcharge(data_dir: Path) -> None:
    tariff_path = data_dir / "tariff.py"
    assert tariff_path.exists()

    spec = importlib.util.spec_from_file_location("tariff_testcopy", str(tariff_path))
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]

    assert hasattr(mod, "ModalSurcharge")
    v = mod.ModalSurcharge(12)
    assert isinstance(v, float)


---

### FILENAME: test_func_parity.py
CONTENT:

# tests/test_func_parity.py
from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Iterable, Set

# VBA: Public if not explicitly Private
VBA_DECL_RE = re.compile(
    r"(?im)^\s*(?:(public|private|friend)\s+)?(function|sub)\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)

# Some VBA members we intentionally ignore (e.g., worksheet event stubs, if any appear)
VBA_IGNORE_NAMES: Set[str] = {
    # add known non-base / event handlers here if they appear in Mod_*.txt
}

# Python helpers inside basfunct.py we don't want to treat as "ported VBA names"
PY_IGNORE_NAMES: Set[str] = {
    "DataRepo",
    "_excel_round",
}


def _find_module_txt_files() -> list[Path]:
    # Prefer project root, but support /mnt/data (LLM sandbox)
    roots = [Path.cwd(), Path("/mnt/data")]
    files: list[Path] = []
    for root in roots:
        files.extend(sorted(root.glob("Mod_*.txt")))
    # de-dupe by resolved path string
    uniq: dict[str, Path] = {}
    for p in files:
        try:
            uniq[str(p.resolve())] = p
        except Exception:
            uniq[str(p)] = p
    return list(uniq.values())


def _vba_public_names_from_text(text: str) -> Set[str]:
    names: Set[str] = set()
    for m in VBA_DECL_RE.finditer(text):
        vis = (m.group(1) or "").strip().lower()
        kind = (m.group(2) or "").strip().lower()
        name = (m.group(3) or "").strip()

        # Only Functions/Subs
        if kind not in ("function", "sub"):
            continue

        # Exclude Private
        if vis == "private":
            continue

        if name in VBA_IGNORE_NAMES:
            continue

        names.add(name)
    return names


def _collect_vba_public_names() -> Set[str]:
    mod_files = _find_module_txt_files()
    assert mod_files, "No Mod_*.txt files found (run TASK 2 VBA export)."

    names: Set[str] = set()
    for p in mod_files:
        text = p.read_text(encoding="utf-8", errors="ignore")
        names |= _vba_public_names_from_text(text)
    assert names, "No public VBA Function/Sub names found in Mod_*.txt files."
    return names


def _collect_python_def_names(basfunct_path: Path) -> Set[str]:
    src = basfunct_path.read_text(encoding="utf-8")
    tree = ast.parse(src, filename=str(basfunct_path))
    names: Set[str] = set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            if node.name in PY_IGNORE_NAMES:
                continue
            if node.name.startswith("_"):
                continue
            names.add(node.name)
    return names


def test_public_vba_names_have_python_defs() -> None:
    vba_names = _collect_vba_public_names()

    # Locate basfunct.py in project root (or /mnt/data fallback)
    basfunct_candidates = [Path.cwd() / "basfunct.py", Path("/mnt/data") / "basfunct.py"]
    basfunct_path = next((p for p in basfunct_candidates if p.exists()), None)
    assert basfunct_path is not None, "basfunct.py not found (create it in TASK 5A)."

    py_names = _collect_python_def_names(basfunct_path)

    missing = sorted(n for n in vba_names if n not in py_names)
    assert not missing, f"Missing Python defs for VBA public names: {missing}"


def test_no_duplicate_python_defs_for_vba_names() -> None:
    """
    Ensures each VBA public name maps to exactly one Python def.
    (In Python, duplicate def names in the same module would overwrite; we detect that by AST only
    yielding final name set. So we do a stricter check by scanning raw text for 'def <name>(' counts.)
    """
    vba_names = _collect_vba_public_names()

    basfunct_candidates = [Path.cwd() / "basfunct.py", Path("/mnt/data") / "basfunct.py"]
    basfunct_path = next((p for p in basfunct_candidates if p.exists()), None)
    assert basfunct_path is not None, "basfunct.py not found (create it in TASK 5A)."

    src = basfunct_path.read_text(encoding="utf-8")

    duplicates = []
    for name in sorted(vba_names):
        # exact 'def Name(' occurrences
        cnt = len(re.findall(rf"(?m)^\s*def\s+{re.escape(name)}\s*\(", src))
        if cnt != 1:
            duplicates.append((name, cnt))

    assert not duplicates, f"Expected exactly one Python def for each VBA name; mismatches: {duplicates}"


---


## 5. Data Artifacts

### FILENAME: limits.csv
CONTENT:

Name,Value
MinAgeFlex,60.0
MinTermFlex,5.0


---

### FILENAME: tables.csv
CONTENT:

Name,Value
0.0|4,"{""x/y"": ""0.0"", ""DAV1994_T_M"": ""0.011687"", ""DAV1994_T_F"": ""0.009003"", ""DAV2008_T_M"": ""0.006113"", ""DAV2008_T_F"": ""0.005088""}"
1.0|5,"{""x/y"": ""1.0"", ""DAV1994_T_M"": ""0.001008"", ""DAV1994_T_F"": ""0.000867"", ""DAV2008_T_M"": ""0.000423"", ""DAV2008_T_F"": ""0.000387""}"
2.0|6,"{""x/y"": ""2.0"", ""DAV1994_T_M"": ""0.000728"", ""DAV1994_T_F"": ""0.000624"", ""DAV2008_T_M"": ""0.000343"", ""DAV2008_T_F"": ""0.000318""}"
3.0|7,"{""x/y"": ""3.0"", ""DAV1994_T_M"": ""0.000542"", ""DAV1994_T_F"": ""0.000444"", ""DAV2008_T_M"": ""0.000275"", ""DAV2008_T_F"": ""0.000255""}"
4.0|8,"{""x/y"": ""4.0"", ""DAV1994_T_M"": ""0.000473"", ""DAV1994_T_F"": ""0.000345"", ""DAV2008_T_M"": ""0.00022"", ""DAV2008_T_F"": ""0.000202""}"
5.0|9,"{""x/y"": ""5.0"", ""DAV1994_T_M"": ""0.000452"", ""DAV1994_T_F"": ""0.000307"", ""DAV2008_T_M"": ""0.000182"", ""DAV2008_T_F"": ""0.000163""}"
6.0|10,"{""x/y"": ""6.0"", ""DAV1994_T_M"": ""0.000433"", ""DAV1994_T_F"": ""0.000293"", ""DAV2008_T_M"": ""0.000155"", ""DAV2008_T_F"": ""0.000134""}"
7.0|11,"{""x/y"": ""7.0"", ""DAV1994_T_M"": ""0.000408"", ""DAV1994_T_F"": ""0.000283"", ""DAV2008_T_M"": ""0.000139"", ""DAV2008_T_F"": ""0.000115""}"
8.0|12,"{""x/y"": ""8.0"", ""DAV1994_T_M"": ""0.000379"", ""DAV1994_T_F"": ""0.000275"", ""DAV2008_T_M"": ""0.000129"", ""DAV2008_T_F"": ""0.000105""}"
9.0|13,"{""x/y"": ""9.0"", ""DAV1994_T_M"": ""0.000352"", ""DAV1994_T_F"": ""0.000268"", ""DAV2008_T_M"": ""0.000125"", ""DAV2008_T_F"": ""9.9e-05""}"
10.0|14,"{""x/y"": ""10.0"", ""DAV1994_T_M"": ""0.000334"", ""DAV1994_T_F"": ""0.000261"", ""DAV2008_T_M"": ""0.000129"", ""DAV2008_T_F"": ""0.000102""}"
11.0|15,"{""x/y"": ""11.0"", ""DAV1994_T_M"": ""0.000331"", ""DAV1994_T_F"": ""0.00026"", ""DAV2008_T_M"": ""0.000143"", ""DAV2008_T_F"": ""0.000111""}"
12.0|16,"{""x/y"": ""12.0"", ""DAV1994_T_M"": ""0.00034"", ""DAV1994_T_F"": ""0.000267"", ""DAV2008_T_M"": ""0.000173"", ""DAV2008_T_F"": ""0.000127""}"
13.0|17,"{""x/y"": ""13.0"", ""DAV1994_T_M"": ""0.000371"", ""DAV1994_T_F"": ""0.000281"", ""DAV2008_T_M"": ""0.000222"", ""DAV2008_T_F"": ""0.000153""}"
14.0|18,"{""x/y"": ""14.0"", ""DAV1994_T_M"": ""0.000451"", ""DAV1994_T_F"": ""0.000307"", ""DAV2008_T_M"": ""0.000303"", ""DAV2008_T_F"": ""0.000188""}"
15.0|19,"{""x/y"": ""15.0"", ""DAV1994_T_M"": ""0.000593"", ""DAV1994_T_F"": ""0.000353"", ""DAV2008_T_M"": ""0.000417"", ""DAV2008_T_F"": ""0.000228""}"
16.0|20,"{""x/y"": ""16.0"", ""DAV1994_T_M"": ""0.000792"", ""DAV1994_T_F"": ""0.000416"", ""DAV2008_T_M"": ""0.000557"", ""DAV2008_T_F"": ""0.000271""}"
17.0|21,"{""x/y"": ""17.0"", ""DAV1994_T_M"": ""0.00104"", ""DAV1994_T_F"": ""0.00048"", ""DAV2008_T_M"": ""0.000709"", ""DAV2008_T_F"": ""0.00031""}"
18.0|22,"{""x/y"": ""18.0"", ""DAV1994_T_M"": ""0.001298"", ""DAV1994_T_F"": ""0.000537"", ""DAV2008_T_M"": ""0.00085"", ""DAV2008_T_F"": ""0.000324""}"
19.0|23,"{""x/y"": ""19.0"", ""DAV1994_T_M"": ""0.001437"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.000953"", ""DAV2008_T_F"": ""0.00033""}"
20.0|24,"{""x/y"": ""20.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.001012"", ""DAV2008_T_F"": ""0.000328""}"
21.0|25,"{""x/y"": ""21.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.001022"", ""DAV2008_T_F"": ""0.000322""}"
22.0|26,"{""x/y"": ""22.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.001004"", ""DAV2008_T_F"": ""0.000314""}"
23.0|27,"{""x/y"": ""23.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.000963"", ""DAV2008_T_F"": ""0.000304""}"
24.0|28,"{""x/y"": ""24.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.000911"", ""DAV2008_T_F"": ""0.000297""}"
25.0|29,"{""x/y"": ""25.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.000856"", ""DAV2008_T_F"": ""0.000293""}"
26.0|30,"{""x/y"": ""26.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.00056"", ""DAV2008_T_M"": ""0.000808"", ""DAV2008_T_F"": ""0.000292""}"
27.0|31,"{""x/y"": ""27.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.000581"", ""DAV2008_T_M"": ""0.000772"", ""DAV2008_T_F"": ""0.000292""}"
28.0|32,"{""x/y"": ""28.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.000612"", ""DAV2008_T_M"": ""0.000752"", ""DAV2008_T_F"": ""0.000296""}"
29.0|33,"{""x/y"": ""29.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.000645"", ""DAV2008_T_M"": ""0.000745"", ""DAV2008_T_F"": ""0.000302""}"
30.0|34,"{""x/y"": ""30.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.000689"", ""DAV2008_T_M"": ""0.000752"", ""DAV2008_T_F"": ""0.000311""}"
31.0|35,"{""x/y"": ""31.0"", ""DAV1994_T_M"": ""0.001476"", ""DAV1994_T_F"": ""0.000735"", ""DAV2008_T_M"": ""0.000768"", ""DAV2008_T_F"": ""0.000327""}"
32.0|36,"{""x/y"": ""32.0"", ""DAV1994_T_M"": ""0.001489"", ""DAV1994_T_F"": ""0.000783"", ""DAV2008_T_M"": ""0.000791"", ""DAV2008_T_F"": ""0.000351""}"
33.0|37,"{""x/y"": ""33.0"", ""DAV1994_T_M"": ""0.001551"", ""DAV1994_T_F"": ""0.000833"", ""DAV2008_T_M"": ""0.00082"", ""DAV2008_T_F"": ""0.000386""}"
34.0|38,"{""x/y"": ""34.0"", ""DAV1994_T_M"": ""0.001641"", ""DAV1994_T_F"": ""0.000897"", ""DAV2008_T_M"": ""0.000855"", ""DAV2008_T_F"": ""0.000433""}"
35.0|39,"{""x/y"": ""35.0"", ""DAV1994_T_M"": ""0.001747"", ""DAV1994_T_F"": ""0.000971"", ""DAV2008_T_M"": ""0.000895"", ""DAV2008_T_F"": ""0.00049""}"
36.0|40,"{""x/y"": ""36.0"", ""DAV1994_T_M"": ""0.001869"", ""DAV1994_T_F"": ""0.001057"", ""DAV2008_T_M"": ""0.000945"", ""DAV2008_T_F"": ""0.000555""}"
37.0|41,"{""x/y"": ""37.0"", ""DAV1994_T_M"": ""0.002007"", ""DAV1994_T_F"": ""0.001156"", ""DAV2008_T_M"": ""0.001005"", ""DAV2008_T_F"": ""0.000624""}"
38.0|42,"{""x/y"": ""38.0"", ""DAV1994_T_M"": ""0.002167"", ""DAV1994_T_F"": ""0.001267"", ""DAV2008_T_M"": ""0.001083"", ""DAV2008_T_F"": ""0.000701""}"
39.0|43,"{""x/y"": ""39.0"", ""DAV1994_T_M"": ""0.002354"", ""DAV1994_T_F"": ""0.00139"", ""DAV2008_T_M"": ""0.001181"", ""DAV2008_T_F"": ""0.000783""}"
40.0|44,"{""x/y"": ""40.0"", ""DAV1994_T_M"": ""0.002569"", ""DAV1994_T_F"": ""0.001524"", ""DAV2008_T_M"": ""0.001301"", ""DAV2008_T_F"": ""0.000872""}"
41.0|45,"{""x/y"": ""41.0"", ""DAV1994_T_M"": ""0.002823"", ""DAV1994_T_F"": ""0.001672"", ""DAV2008_T_M"": ""0.001447"", ""DAV2008_T_F"": ""0.000972""}"
42.0|46,"{""x/y"": ""42.0"", ""DAV1994_T_M"": ""0.003087"", ""DAV1994_T_F"": ""0.001812"", ""DAV2008_T_M"": ""0.001623"", ""DAV2008_T_F"": ""0.001084""}"
43.0|47,"{""x/y"": ""43.0"", ""DAV1994_T_M"": ""0.003387"", ""DAV1994_T_F"": ""0.001964"", ""DAV2008_T_M"": ""0.001833"", ""DAV2008_T_F"": ""0.001213""}"
44.0|48,"{""x/y"": ""44.0"", ""DAV1994_T_M"": ""0.003726"", ""DAV1994_T_F"": ""0.002126"", ""DAV2008_T_M"": ""0.002082"", ""DAV2008_T_F"": ""0.001359""}"
45.0|49,"{""x/y"": ""45.0"", ""DAV1994_T_M"": ""0.0041"", ""DAV1994_T_F"": ""0.002295"", ""DAV2008_T_M"": ""0.002364"", ""DAV2008_T_F"": ""0.001524""}"
46.0|50,"{""x/y"": ""46.0"", ""DAV1994_T_M"": ""0.004522"", ""DAV1994_T_F"": ""0.00248"", ""DAV2008_T_M"": ""0.002669"", ""DAV2008_T_F"": ""0.001706""}"
47.0|51,"{""x/y"": ""47.0"", ""DAV1994_T_M"": ""0.004983"", ""DAV1994_T_F"": ""0.002676"", ""DAV2008_T_M"": ""0.002983"", ""DAV2008_T_F"": ""0.001903""}"
48.0|52,"{""x/y"": ""48.0"", ""DAV1994_T_M"": ""0.005508"", ""DAV1994_T_F"": ""0.002902"", ""DAV2008_T_M"": ""0.003302"", ""DAV2008_T_F"": ""0.002109""}"
49.0|53,"{""x/y"": ""49.0"", ""DAV1994_T_M"": ""0.006094"", ""DAV1994_T_F"": ""0.003151"", ""DAV2008_T_M"": ""0.00363"", ""DAV2008_T_F"": ""0.002324""}"
50.0|54,"{""x/y"": ""50.0"", ""DAV1994_T_M"": ""0.006751"", ""DAV1994_T_F"": ""0.003425"", ""DAV2008_T_M"": ""0.003981"", ""DAV2008_T_F"": ""0.002546""}"
51.0|55,"{""x/y"": ""51.0"", ""DAV1994_T_M"": ""0.007485"", ""DAV1994_T_F"": ""0.003728"", ""DAV2008_T_M"": ""0.004371"", ""DAV2008_T_F"": ""0.002782""}"
52.0|56,"{""x/y"": ""52.0"", ""DAV1994_T_M"": ""0.008302"", ""DAV1994_T_F"": ""0.004066"", ""DAV2008_T_M"": ""0.004812"", ""DAV2008_T_F"": ""0.003035""}"
53.0|57,"{""x/y"": ""53.0"", ""DAV1994_T_M"": ""0.009215"", ""DAV1994_T_F"": ""0.00445"", ""DAV2008_T_M"": ""0.005308"", ""DAV2008_T_F"": ""0.003306""}"
54.0|58,"{""x/y"": ""54.0"", ""DAV1994_T_M"": ""0.010195"", ""DAV1994_T_F"": ""0.004862"", ""DAV2008_T_M"": ""0.005857"", ""DAV2008_T_F"": ""0.003593""}"
55.0|59,"{""x/y"": ""55.0"", ""DAV1994_T_M"": ""0.011236"", ""DAV1994_T_F"": ""0.005303"", ""DAV2008_T_M"": ""0.00646"", ""DAV2008_T_F"": ""0.003898""}"
56.0|60,"{""x/y"": ""56.0"", ""DAV1994_T_M"": ""0.01234"", ""DAV1994_T_F"": ""0.005777"", ""DAV2008_T_M"": ""0.007117"", ""DAV2008_T_F"": ""0.004228""}"
57.0|61,"{""x/y"": ""57.0"", ""DAV1994_T_M"": ""0.013519"", ""DAV1994_T_F"": ""0.006302"", ""DAV2008_T_M"": ""0.007831"", ""DAV2008_T_F"": ""0.004585""}"
58.0|62,"{""x/y"": ""58.0"", ""DAV1994_T_M"": ""0.014784"", ""DAV1994_T_F"": ""0.006884"", ""DAV2008_T_M"": ""0.008604"", ""DAV2008_T_F"": ""0.004974""}"
59.0|63,"{""x/y"": ""59.0"", ""DAV1994_T_M"": ""0.01615"", ""DAV1994_T_F"": ""0.00753"", ""DAV2008_T_M"": ""0.009454"", ""DAV2008_T_F"": ""0.005402""}"
60.0|64,"{""x/y"": ""60.0"", ""DAV1994_T_M"": ""0.017625"", ""DAV1994_T_F"": ""0.00824"", ""DAV2008_T_M"": ""0.010404"", ""DAV2008_T_F"": ""0.005884""}"
61.0|65,"{""x/y"": ""61.0"", ""DAV1994_T_M"": ""0.019223"", ""DAV1994_T_F"": ""0.009022"", ""DAV2008_T_M"": ""0.011504"", ""DAV2008_T_F"": ""0.006449""}"
62.0|66,"{""x/y"": ""62.0"", ""DAV1994_T_M"": ""0.020956"", ""DAV1994_T_F"": ""0.009884"", ""DAV2008_T_M"": ""0.012818"", ""DAV2008_T_F"": ""0.007126""}"
63.0|67,"{""x/y"": ""63.0"", ""DAV1994_T_M"": ""0.022833"", ""DAV1994_T_F"": ""0.010839"", ""DAV2008_T_M"": ""0.014429"", ""DAV2008_T_F"": ""0.007935""}"
64.0|68,"{""x/y"": ""64.0"", ""DAV1994_T_M"": ""0.024858"", ""DAV1994_T_F"": ""0.011889"", ""DAV2008_T_M"": ""0.016415"", ""DAV2008_T_F"": ""0.008898""}"
65.0|69,"{""x/y"": ""65.0"", ""DAV1994_T_M"": ""0.027073"", ""DAV1994_T_F"": ""0.013054"", ""DAV2008_T_M"": ""0.018832"", ""DAV2008_T_F"": ""0.010025""}"
66.0|70,"{""x/y"": ""66.0"", ""DAV1994_T_M"": ""0.029552"", ""DAV1994_T_F"": ""0.014371"", ""DAV2008_T_M"": ""0.021704"", ""DAV2008_T_F"": ""0.011323""}"
67.0|71,"{""x/y"": ""67.0"", ""DAV1994_T_M"": ""0.03235"", ""DAV1994_T_F"": ""0.015874"", ""DAV2008_T_M"": ""0.025016"", ""DAV2008_T_F"": ""0.012797""}"
68.0|72,"{""x/y"": ""68.0"", ""DAV1994_T_M"": ""0.035632"", ""DAV1994_T_F"": ""0.017667"", ""DAV2008_T_M"": ""0.028738"", ""DAV2008_T_F"": ""0.01446""}"
69.0|73,"{""x/y"": ""69.0"", ""DAV1994_T_M"": ""0.039224"", ""DAV1994_T_F"": ""0.019657"", ""DAV2008_T_M"": ""0.032822"", ""DAV2008_T_F"": ""0.016332""}"
70.0|74,"{""x/y"": ""70.0"", ""DAV1994_T_M"": ""0.043127"", ""DAV1994_T_F"": ""0.021861"", ""DAV2008_T_M"": ""0.037219"", ""DAV2008_T_F"": ""0.01844""}"
71.0|75,"{""x/y"": ""71.0"", ""DAV1994_T_M"": ""0.0474"", ""DAV1994_T_F"": ""0.024344"", ""DAV2008_T_M"": ""0.04188"", ""DAV2008_T_F"": ""0.020813""}"
72.0|76,"{""x/y"": ""72.0"", ""DAV1994_T_M"": ""0.05211"", ""DAV1994_T_F"": ""0.027191"", ""DAV2008_T_M"": ""0.046597"", ""DAV2008_T_F"": ""0.023475""}"
73.0|77,"{""x/y"": ""73.0"", ""DAV1994_T_M"": ""0.057472"", ""DAV1994_T_F"": ""0.030576"", ""DAV2008_T_M"": ""0.051181"", ""DAV2008_T_F"": ""0.027035""}"
74.0|78,"{""x/y"": ""74.0"", ""DAV1994_T_M"": ""0.06344"", ""DAV1994_T_F"": ""0.034504"", ""DAV2008_T_M"": ""0.05611"", ""DAV2008_T_F"": ""0.030413""}"
75.0|79,"{""x/y"": ""75.0"", ""DAV1994_T_M"": ""0.070039"", ""DAV1994_T_F"": ""0.03903"", ""DAV2008_T_M"": ""0.061477"", ""DAV2008_T_F"": ""0.034287""}"
76.0|80,"{""x/y"": ""76.0"", ""DAV1994_T_M"": ""0.077248"", ""DAV1994_T_F"": ""0.044184"", ""DAV2008_T_M"": ""0.067433"", ""DAV2008_T_F"": ""0.038749""}"
77.0|81,"{""x/y"": ""77.0"", ""DAV1994_T_M"": ""0.085073"", ""DAV1994_T_F"": ""0.050014"", ""DAV2008_T_M"": ""0.07416"", ""DAV2008_T_F"": ""0.043937""}"
78.0|82,"{""x/y"": ""78.0"", ""DAV1994_T_M"": ""0.093534"", ""DAV1994_T_F"": ""0.056574"", ""DAV2008_T_M"": ""0.081806"", ""DAV2008_T_F"": ""0.049993""}"
79.0|83,"{""x/y"": ""79.0"", ""DAV1994_T_M"": ""0.102662"", ""DAV1994_T_F"": ""0.063921"", ""DAV2008_T_M"": ""0.090478"", ""DAV2008_T_F"": ""0.057024""}"
80.0|84,"{""x/y"": ""80.0"", ""DAV1994_T_M"": ""0.112477"", ""DAV1994_T_F"": ""0.072101"", ""DAV2008_T_M"": ""0.100261"", ""DAV2008_T_F"": ""0.065113""}"
81.0|85,"{""x/y"": ""81.0"", ""DAV1994_T_M"": ""0.122995"", ""DAV1994_T_F"": ""0.081151"", ""DAV2008_T_M"": ""0.111193"", ""DAV2008_T_F"": ""0.074288""}"
82.0|86,"{""x/y"": ""82.0"", ""DAV1994_T_M"": ""0.134231"", ""DAV1994_T_F"": ""0.091096"", ""DAV2008_T_M"": ""0.123283"", ""DAV2008_T_F"": ""0.08459""}"
83.0|87,"{""x/y"": ""83.0"", ""DAV1994_T_M"": ""0.146212"", ""DAV1994_T_F"": ""0.10197"", ""DAV2008_T_M"": ""0.136498"", ""DAV2008_T_F"": ""0.096095""}"
84.0|88,"{""x/y"": ""84.0"", ""DAV1994_T_M"": ""0.158964"", ""DAV1994_T_F"": ""0.113798"", ""DAV2008_T_M"": ""0.150887"", ""DAV2008_T_F"": ""0.109028""}"
85.0|89,"{""x/y"": ""85.0"", ""DAV1994_T_M"": ""0.172512"", ""DAV1994_T_F"": ""0.126628"", ""DAV2008_T_M"": ""0.1665"", ""DAV2008_T_F"": ""0.123611""}"
86.0|90,"{""x/y"": ""86.0"", ""DAV1994_T_M"": ""0.186896"", ""DAV1994_T_F"": ""0.140479"", ""DAV2008_T_M"": ""0.183344"", ""DAV2008_T_F"": ""0.140022""}"
87.0|91,"{""x/y"": ""87.0"", ""DAV1994_T_M"": ""0.202185"", ""DAV1994_T_F"": ""0.155379"", ""DAV2008_T_M"": ""0.201323"", ""DAV2008_T_F"": ""0.158257""}"
88.0|92,"{""x/y"": ""88.0"", ""DAV1994_T_M"": ""0.218413"", ""DAV1994_T_F"": ""0.171325"", ""DAV2008_T_M"": ""0.220284"", ""DAV2008_T_F"": ""0.178185""}"
89.0|93,"{""x/y"": ""89.0"", ""DAV1994_T_M"": ""0.235597"", ""DAV1994_T_F"": ""0.188318"", ""DAV2008_T_M"": ""0.240073"", ""DAV2008_T_F"": ""0.199669""}"
90.0|94,"{""x/y"": ""90.0"", ""DAV1994_T_M"": ""0.253691"", ""DAV1994_T_F"": ""0.206375"", ""DAV2008_T_M"": ""0.260556"", ""DAV2008_T_F"": ""0.222504""}"
91.0|95,"{""x/y"": ""91.0"", ""DAV1994_T_M"": ""0.272891"", ""DAV1994_T_F"": ""0.225558"", ""DAV2008_T_M"": ""0.281602"", ""DAV2008_T_F"": ""0.246453""}"
92.0|96,"{""x/y"": ""92.0"", ""DAV1994_T_M"": ""0.293142"", ""DAV1994_T_F"": ""0.245839"", ""DAV2008_T_M"": ""0.303079"", ""DAV2008_T_F"": ""0.271195""}"
93.0|97,"{""x/y"": ""93.0"", ""DAV1994_T_M"": ""0.314638"", ""DAV1994_T_F"": ""0.26727"", ""DAV2008_T_M"": ""0.324872"", ""DAV2008_T_F"": ""0.295584""}"
94.0|98,"{""x/y"": ""94.0"", ""DAV1994_T_M"": ""0.337739"", ""DAV1994_T_F"": ""0.289983"", ""DAV2008_T_M"": ""0.346887"", ""DAV2008_T_F"": ""0.319362""}"
95.0|99,"{""x/y"": ""95.0"", ""DAV1994_T_M"": ""0.36206"", ""DAV1994_T_F"": ""0.314007"", ""DAV2008_T_M"": ""0.369051"", ""DAV2008_T_F"": ""0.343441""}"
96.0|100,"{""x/y"": ""96.0"", ""DAV1994_T_M"": ""0.388732"", ""DAV1994_T_F"": ""0.340119"", ""DAV2008_T_M"": ""0.391305"", ""DAV2008_T_F"": ""0.367818""}"
97.0|101,"{""x/y"": ""97.0"", ""DAV1994_T_M"": ""0.419166"", ""DAV1994_T_F"": ""0.367388"", ""DAV2008_T_M"": ""0.413938"", ""DAV2008_T_F"": ""0.392493""}"
98.0|102,"{""x/y"": ""98.0"", ""DAV1994_T_M"": ""0.452008"", ""DAV1994_T_F"": ""0.397027"", ""DAV2008_T_M"": ""0.437313"", ""DAV2008_T_F"": ""0.41746""}"
99.0|103,"{""x/y"": ""99.0"", ""DAV1994_T_M"": ""0.4864"", ""DAV1994_T_F"": ""0.428748"", ""DAV2008_T_M"": ""0.461101"", ""DAV2008_T_F"": ""0.442716""}"
100.0|104,"{""x/y"": ""100.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.485304"", ""DAV2008_T_F"": ""0.468258""}"
101.0|105,"{""x/y"": ""101.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.509924"", ""DAV2008_T_F"": ""0.494075""}"
102.0|106,"{""x/y"": ""102.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.534957"", ""DAV2008_T_F"": ""0.520164""}"
103.0|107,"{""x/y"": ""103.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.560407"", ""DAV2008_T_F"": ""0.546514""}"
104.0|108,"{""x/y"": ""104.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.586265"", ""DAV2008_T_F"": ""0.573114""}"
105.0|109,"{""x/y"": ""105.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.612529"", ""DAV2008_T_F"": ""0.599953""}"
106.0|110,"{""x/y"": ""106.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.639188"", ""DAV2008_T_F"": ""0.627014""}"
107.0|111,"{""x/y"": ""107.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.666233"", ""DAV2008_T_F"": ""0.654283""}"
108.0|112,"{""x/y"": ""108.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.693651"", ""DAV2008_T_F"": ""0.681741""}"
109.0|113,"{""x/y"": ""109.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.721425"", ""DAV2008_T_F"": ""0.709364""}"
110.0|114,"{""x/y"": ""110.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.749533"", ""DAV2008_T_F"": ""0.73713""}"
111.0|115,"{""x/y"": ""111.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.77795"", ""DAV2008_T_F"": ""0.765011""}"
112.0|116,"{""x/y"": ""112.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.806647"", ""DAV2008_T_F"": ""0.792974""}"
113.0|117,"{""x/y"": ""113.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.835585"", ""DAV2008_T_F"": ""0.820987""}"
114.0|118,"{""x/y"": ""114.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.864722"", ""DAV2008_T_F"": ""0.849009""}"
115.0|119,"{""x/y"": ""115.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.894008"", ""DAV2008_T_F"": ""0.876998""}"
116.0|120,"{""x/y"": ""116.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.923382"", ""DAV2008_T_F"": ""0.904905""}"
117.0|121,"{""x/y"": ""117.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.952778"", ""DAV2008_T_F"": ""0.932675""}"
118.0|122,"{""x/y"": ""118.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""0.982113"", ""DAV2008_T_F"": ""0.960249""}"
119.0|123,"{""x/y"": ""119.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""1.0"", ""DAV2008_T_F"": ""0.987564""}"
120.0|124,"{""x/y"": ""120.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""1.0"", ""DAV2008_T_F"": ""1.0""}"
121.0|125,"{""x/y"": ""121.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""1.0"", ""DAV2008_T_F"": ""1.0""}"
122.0|126,"{""x/y"": ""122.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""1.0"", ""DAV2008_T_F"": ""1.0""}"
123.0|127,"{""x/y"": ""123.0"", ""DAV1994_T_M"": ""1.0"", ""DAV1994_T_F"": ""1.0"", ""DAV2008_T_M"": ""1.0"", ""DAV2008_T_F"": ""1.0""}"


---

### FILENAME: tariff.csv
CONTENT:

Name,Value
InterestRate,0.0175
MortalityTable,DAV1994_T
alpha,0.025
beta1,0.025
gamma1,0.0008
gamma2,0.00125
gamma3,0.0025
k,24.0


---

### FILENAME: var.csv
CONTENT:


Name,Value
x,40.0
Sex,M
n,30.0
t,20.0
SumInsured,"""100000"""
PayFreq,12.0


---

