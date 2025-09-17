'''////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>   445 Length Waybill class </summary>
'''
''' <remarks>   Michael Sanders, 10/30/2020. </remarks>
'''////////////////////////////////////////////////////////////////////////////////////////////////////

Public Class Class_445_Waybill

    Private mserial_no As Decimal
    Private mwb_num As Decimal
    Private mwb_date As Date
    Private macct_period As String
    Private mu_cars As Integer
    Private mu_car_init As String
    Private mu_car_num As Decimal
    Private mtofc_serv_code As String
    Private mu_tc_units As Integer
    Private mu_tc_init As String
    Private mu_tc_num As Decimal
    Private mstcc_w49 As String
    Private mbill_wght As Decimal
    Private mact_wght As Decimal
    Private mu_rev As Decimal
    Private mtran_chrg As Decimal
    Private mmisc_chrg As Decimal
    Private mintra_state_code As Byte
    Private mtransit_code As Byte
    Private mall_rail_code As Byte
    Private mtype_move As Byte
    Private mmove_via_water As Byte
    Private mtruck_for_rail As Byte
    Private mshortline_miles As Integer
    Private mrebill As Byte
    Private mstratum As Byte
    Private msubsample As Byte
    Private mtransborder_flg As String
    Private mrate_flg As Byte
    Private mwb_id As String
    Private mreport_rr As Integer
    Private mo_fsac As Decimal
    Private morr As Integer
    Private mjct1 As String
    Private mjrr1 As Integer
    Private mjct2 As String
    Private mjrr2 As Integer
    Private mjct3 As String
    Private mjrr3 As Integer
    Private mjct4 As String
    Private mjrr4 As Integer
    Private mjct5 As String
    Private mjrr5 As Integer
    Private mjct6 As String
    Private mjrr6 As Integer
    Private mjct7 As String
    Private mjrr7 As Integer
    Private mjct8 As String
    Private mjrr8 As Integer
    Private mjct9 As String
    Private mtrr As Integer
    Private mt_fsac As Decimal
    Private mpop_cnt As Decimal
    Private mstratum_cnt As Decimal
    Private mreport_period As Byte
    Private mcar_own_mark As String
    Private mcar_lessee_mark As String
    Private mcar_cap As Decimal
    Private mnom_car_cap As Integer
    Private mtare As Integer
    Private moutside_l As Decimal
    Private moutside_w As Integer
    Private moutside_h As Integer
    Private mex_outside_h As Integer
    Private mtype_wheel As String
    Private mno_axles As String
    Private mdraft_gear As Byte
    Private mart_units As Byte
    Private mpool_code As Decimal
    Private mcar_typ As String
    Private mmech As String
    Private mlic_st As String
    Private mmx_wght_rail As Integer
    Private mo_splc As Decimal
    Private mt_splc As Decimal
    Private mu_fuel_surchrg
    Private merr_code1 As Byte
    Private merr_code2 As Byte
    Private merr_code3 As Byte
    Private merr_code4 As Byte
    Private merr_code5 As Byte
    Private merr_code6 As Byte
    Private merr_code7 As Byte
    Private merr_code8 As Byte
    Private merr_code9 As Byte
    Private merr_code10 As Byte
    Private merr_code11 As Byte
    Private merr_code12 As Byte
    Private merr_code13 As Byte
    Private merr_code14 As Byte
    Private merr_code15 As Byte
    Private merr_code16 As Byte
    Private merr_code17 As Byte
    Private mcar_own As String
    Private mtofc_unit_type As String
    Private malk_flg As Byte
    Private mTracking_No As Long

    Public Property Serial_No() As Decimal
        Get
            Return mserial_no
        End Get
        Set(ByVal value As Decimal)
            mserial_no = value
        End Set
    End Property

    Public Property WB_Num() As Decimal
        Get
            Return mwb_num
        End Get
        Set(ByVal value As Decimal)
            mwb_num = value
        End Set
    End Property

    Public Property WB_Date() As Date
        Get
            Return mwb_date
        End Get
        Set(ByVal value As Date)
            mwb_date = value
        End Set
    End Property

    Public Property Acct_Period() As String
        Get
            Return macct_period
        End Get
        Set(ByVal value As String)
            macct_period = value
        End Set
    End Property

    Public Property U_Cars() As Integer
        Get
            Return mu_cars
        End Get
        Set(ByVal value As Integer)
            mu_cars = value
        End Set
    End Property

    Public Property U_Car_Init() As String
        Get
            Return mu_car_init
        End Get
        Set(ByVal value As String)
            mu_car_init = value
        End Set
    End Property

    Public Property U_Car_Num() As Decimal
        Get
            Return mu_car_num
        End Get
        Set(ByVal value As Decimal)
            mu_car_num = value
        End Set
    End Property

    Public Property TOFC_Serv_Code() As String
        Get
            Return mtofc_serv_code
        End Get
        Set(ByVal value As String)
            mtofc_serv_code = value
        End Set
    End Property

    Public Property U_TC_Units() As Integer
        Get
            Return mu_tc_units
        End Get
        Set(ByVal value As Integer)
            mu_tc_units = value
        End Set
    End Property

    Public Property U_TC_Init() As String
        Get
            Return mu_tc_init
        End Get
        Set(ByVal value As String)
            mu_tc_init = value
        End Set
    End Property

    Public Property U_TC_Num() As Decimal
        Get
            Return mu_tc_num
        End Get
        Set(ByVal value As Decimal)
            mu_tc_num = value
        End Set
    End Property

    Public Property STCC_W49() As String
        Get
            Return mstcc_w49
        End Get
        Set(ByVal value As String)
            mstcc_w49 = value
        End Set
    End Property

    Public Property Bill_Wght() As Decimal
        Get
            Return mbill_wght
        End Get
        Set(ByVal value As Decimal)
            mbill_wght = value
        End Set
    End Property

    Public Property Act_Wght() As Decimal
        Get
            Return mact_wght
        End Get
        Set(ByVal value As Decimal)
            mact_wght = value
        End Set
    End Property

    Public Property U_Rev() As Decimal
        Get
            Return mu_rev
        End Get
        Set(ByVal value As Decimal)
            mu_rev = value
        End Set
    End Property

    Public Property Tran_Chrg() As Decimal
        Get
            Return mtran_chrg
        End Get
        Set(ByVal value As Decimal)
            mtran_chrg = value
        End Set
    End Property

    Public Property Misc_Chrg() As Decimal
        Get
            Return mmisc_chrg
        End Get
        Set(ByVal value As Decimal)
            mmisc_chrg = value
        End Set
    End Property

    Public Property Intra_State_Code() As Byte
        Get
            Return mintra_state_code
        End Get
        Set(ByVal value As Byte)
            mintra_state_code = value
        End Set
    End Property

    Public Property Transit_Code() As Byte
        Get
            Return mtransit_code
        End Get
        Set(ByVal value As Byte)
            mtransit_code = value
        End Set
    End Property

    Public Property All_Rail_Code() As Byte
        Get
            Return mall_rail_code
        End Get
        Set(ByVal value As Byte)
            mall_rail_code = value
        End Set
    End Property

    Public Property Type_Move() As Byte
        Get
            Return mtype_move
        End Get
        Set(ByVal value As Byte)
            mtype_move = value
        End Set
    End Property

    Public Property Move_Via_Water() As Byte
        Get
            Return mmove_via_water
        End Get
        Set(ByVal value As Byte)
            mmove_via_water = value
        End Set
    End Property

    Public Property Truck_For_Rail() As Byte
        Get
            Return mtruck_for_rail
        End Get
        Set(ByVal value As Byte)
            mtruck_for_rail = value
        End Set
    End Property

    Public Property Shortline_Miles() As Integer
        Get
            Return mshortline_miles
        End Get
        Set(ByVal value As Integer)
            mshortline_miles = value
        End Set
    End Property

    Public Property Rebill() As Byte
        Get
            Return mrebill
        End Get
        Set(ByVal value As Byte)
            mrebill = value
        End Set
    End Property

    Public Property Stratum() As Byte
        Get
            Return mstratum
        End Get
        Set(ByVal value As Byte)
            mstratum = value
        End Set
    End Property

    Public Property Subsample() As Byte
        Get
            Return msubsample
        End Get
        Set(ByVal value As Byte)
            msubsample = value
        End Set
    End Property

    Public Property Rate_Flg() As Byte
        Get
            Return mrate_flg
        End Get
        Set(ByVal value As Byte)
            mrate_flg = value
        End Set
    End Property

    Public Property Wb_Id() As String
        Get
            Return mwb_id
        End Get
        Set(ByVal value As String)
            mwb_id = value
        End Set
    End Property

    Public Property Report_RR() As Integer
        Get
            Return mreport_rr
        End Get
        Set(ByVal value As Integer)
            mreport_rr = value
        End Set
    End Property

    Public Property O_FSAC() As Decimal
        Get
            Return mo_fsac
        End Get
        Set(ByVal value As Decimal)
            mo_fsac = value
        End Set
    End Property

    Public Property ORR() As Integer
        Get
            Return morr
        End Get
        Set(ByVal value As Integer)
            morr = value
        End Set
    End Property

    Public Property JCT1() As String
        Get
            Return mjct1
        End Get
        Set(ByVal value As String)
            mjct1 = value
        End Set
    End Property

    Public Property JRR1() As Integer
        Get
            Return mjrr1
        End Get
        Set(ByVal value As Integer)
            mjrr1 = value
        End Set
    End Property

    Public Property JCT2() As String
        Get
            Return mjct2
        End Get
        Set(ByVal value As String)
            mjct2 = value
        End Set
    End Property

    Public Property JRR2() As Integer
        Get
            Return mjrr2
        End Get
        Set(ByVal value As Integer)
            mjrr2 = value
        End Set
    End Property

    Public Property JCT3() As String
        Get
            Return mjct3
        End Get
        Set(ByVal value As String)
            mjct3 = value
        End Set
    End Property

    Public Property JRR3() As Integer
        Get
            Return mjrr3
        End Get
        Set(ByVal value As Integer)
            mjrr3 = value
        End Set
    End Property

    Public Property JCT4() As String
        Get
            Return mjct4
        End Get
        Set(ByVal value As String)
            mjct4 = value
        End Set
    End Property

    Public Property JRR4() As Integer
        Get
            Return mjrr4
        End Get
        Set(ByVal value As Integer)
            mjrr4 = value
        End Set
    End Property

    Public Property JCT5() As String
        Get
            Return mjct5
        End Get
        Set(ByVal value As String)
            mjct5 = value
        End Set
    End Property

    Public Property JRR5() As Integer
        Get
            Return mjrr5
        End Get
        Set(ByVal value As Integer)
            mjrr5 = value
        End Set
    End Property

    Public Property JCT6() As String
        Get
            Return mjct6
        End Get
        Set(ByVal value As String)
            mjct6 = value
        End Set
    End Property

    Public Property JRR6() As Integer
        Get
            Return mjrr6
        End Get
        Set(ByVal value As Integer)
            mjrr6 = value
        End Set
    End Property

    Public Property JCT7() As String
        Get
            Return mjct7
        End Get
        Set(ByVal value As String)
            mjct7 = value
        End Set
    End Property

    Public Property TRR() As Integer
        Get
            Return mtrr
        End Get
        Set(ByVal value As Integer)
            mtrr = value
        End Set
    End Property

    Public Property T_FSAC() As Decimal
        Get
            Return mt_fsac
        End Get
        Set(ByVal value As Decimal)
            mt_fsac = value
        End Set
    End Property

    Public Property Pop_Cnt() As Decimal
        Get
            Return mpop_cnt
        End Get
        Set(ByVal value As Decimal)
            mpop_cnt = value
        End Set
    End Property

    Public Property Stratum_Cnt() As Decimal
        Get
            Return mstratum_cnt
        End Get
        Set(ByVal value As Decimal)
            mstratum_cnt = value
        End Set
    End Property

    Public Property Report_Period() As Byte
        Get
            Return mreport_period
        End Get
        Set(ByVal value As Byte)
            mreport_period = value
        End Set
    End Property

    Public Property Car_Own_Mark() As String
        Get
            Return mcar_own_mark
        End Get
        Set(ByVal value As String)
            mcar_own_mark = value
        End Set
    End Property

    Public Property Car_Lessee_Mark() As String
        Get
            Return mcar_lessee_mark
        End Get
        Set(ByVal value As String)
            mcar_lessee_mark = value
        End Set
    End Property

    Public Property Car_Cap() As Decimal
        Get
            Return mcar_cap
        End Get
        Set(ByVal value As Decimal)
            mcar_cap = value
        End Set
    End Property

    Public Property Nom_Car_Cap() As Integer
        Get
            Return mnom_car_cap
        End Get
        Set(ByVal value As Integer)
            mnom_car_cap = value
        End Set
    End Property

    Public Property Tare() As Integer
        Get
            Return mtare
        End Get
        Set(ByVal value As Integer)
            mtare = value
        End Set
    End Property

    Public Property Outside_L() As Decimal
        Get
            Return moutside_l
        End Get
        Set(ByVal value As Decimal)
            moutside_l = value
        End Set
    End Property

    Public Property Outside_W() As Integer
        Get
            Return moutside_w
        End Get
        Set(ByVal value As Integer)
            moutside_w = value
        End Set
    End Property

    Public Property Outside_H() As Integer
        Get
            Return moutside_h
        End Get
        Set(ByVal value As Integer)
            moutside_h = value
        End Set
    End Property

    Public Property Ex_Outside_H() As Integer
        Get
            Return mex_outside_h
        End Get
        Set(ByVal value As Integer)
            mex_outside_h = value
        End Set
    End Property

    Public Property Type_Wheel() As String
        Get
            Return mtype_wheel
        End Get
        Set(ByVal value As String)
            mtype_wheel = value
        End Set
    End Property

    Public Property No_Axles() As String
        Get
            Return mno_axles
        End Get
        Set(ByVal value As String)
            mno_axles = value
        End Set
    End Property

    Public Property Draft_Gear() As Byte
        Get
            Return mdraft_gear
        End Get
        Set(ByVal value As Byte)
            mdraft_gear = value
        End Set
    End Property

    Public Property Art_Units() As Byte
        Get
            Return mart_units
        End Get
        Set(ByVal value As Byte)
            mart_units = value
        End Set
    End Property

    Public Property Pool_Code() As Decimal
        Get
            Return mpool_code
        End Get
        Set(ByVal value As Decimal)
            mpool_code = value
        End Set
    End Property

    Public Property Car_Typ() As String
        Get
            Return mcar_typ
        End Get
        Set(ByVal value As String)
            mcar_typ = value
        End Set
    End Property

    Public Property Mech() As String
        Get
            Return mmech
        End Get
        Set(ByVal value As String)
            mmech = value
        End Set
    End Property

    Public Property Lic_St() As String
        Get
            Return mlic_st
        End Get
        Set(ByVal value As String)
            mlic_st = value
        End Set
    End Property

    Public Property Mx_Wght_Rail() As Integer
        Get
            Return mmx_wght_rail
        End Get
        Set(ByVal value As Integer)
            mmx_wght_rail = value
        End Set
    End Property

    Public Property O_SPLC() As Decimal
        Get
            Return mo_splc
        End Get
        Set(ByVal value As Decimal)
            mo_splc = value
        End Set
    End Property

    Public Property T_SPLC() As Decimal
        Get
            Return mt_splc
        End Get
        Set(ByVal value As Decimal)
            mt_splc = value
        End Set
    End Property

    Public Property Err_Code1() As Byte
        Get
            Return merr_code1
        End Get
        Set(ByVal value As Byte)
            merr_code1 = value
        End Set
    End Property

    Public Property Err_Code2() As Byte
        Get
            Return merr_code2
        End Get
        Set(ByVal value As Byte)
            merr_code2 = value
        End Set
    End Property

    Public Property Err_Code3() As Byte
        Get
            Return merr_code3
        End Get
        Set(ByVal value As Byte)
            merr_code3 = value
        End Set
    End Property

    Public Property Car_Own() As String
        Get
            Return mcar_own
        End Get
        Set(ByVal value As String)
            mcar_own = value
        End Set
    End Property

    Public Property TOFC_Unit_Type() As String
        Get
            Return mtofc_unit_type
        End Get
        Set(ByVal value As String)
            mtofc_unit_type = value
        End Set
    End Property

    Public Property Transborder_Flg() As String
        Get
            Return mtransborder_flg
        End Get
        Set(ByVal value As String)
            mtransborder_flg = value
        End Set
    End Property

    Public Property U_Fuel_SurChrg() As Decimal
        Get
            Return mu_fuel_surchrg
        End Get
        Set(ByVal value As Decimal)
            mu_fuel_surchrg = value
        End Set
    End Property

    Public Property Tracking_No() As Long
        Get
            Return mTracking_No
        End Get
        Set(ByVal value As Long)
            mTracking_No = value
        End Set
    End Property
End Class
