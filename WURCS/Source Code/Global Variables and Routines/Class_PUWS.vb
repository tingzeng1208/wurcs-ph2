'''////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>   PUWS Waybill class </summary>
'''
''' <remarks>   Michael Sanders, 10/30/2020. </remarks>
'''////////////////////////////////////////////////////////////////////////////////////////////////////
Public Class Class_PUWS

    Private mwb_date As String
    Private macct_period As String
    Private mu_cars As Integer
    Private mcar_own As String
    Private mcar_typ As String
    Private mmech As String
    Private mstb_car_typ As Integer
    Private mtofc_serv_code As String
    Private mu_tc_units As Integer
    Private mtofc_own As String
    Private mtofc_unit_type As String
    Private mhazmat_boxcar As String
    Private mstcc As String
    Private mbill_wght_tons As Integer
    Private mact_wght As Integer
    Private mu_rev As Integer
    Private mtran_chrg As Integer
    Private mmisc_chrg As Integer
    Private mintra_state_code As Integer
    Private mtype_move As Integer
    Private mall_rail_code As Integer
    Private mmove_via_water As Integer
    Private mtransit_code As Integer
    Private mtruck_for_rail As Integer
    Private mrebill As Integer
    Private mshortline_miles As Integer
    Private mstratum As Integer
    Private msubsample As Integer
    Private mexp_factor As Integer
    Private mexp_factor_th As Integer
    Private mjf As Integer
    Private mo_bea As Integer
    Private mo_ft As Integer
    Private mo_st As String
    Private mjct1_st As String
    Private mjct2_st As String
    Private mjct3_st As String
    Private mjct4_st As String
    Private mjct5_st As String
    Private mjct6_st As String
    Private mjct7_st As String
    Private mjct8_st As String
    Private mjct9_st As String
    Private mt_ft As Integer
    Private mt_bea As Integer
    Private mreport_period As Integer
    Private mcar_cap As Integer
    Private mnom_car_cap As Integer
    Private mtare As Integer
    Private moutside_l As Integer
    Private moutside_w As Integer
    Private moutside_h As Integer
    Private mex_outside_h As Integer
    Private mtype_wheel As String
    Private mno_axles As String
    Private mdraft_gear As Integer
    Private mart_units As Integer
    Private merr_code1 As Integer
    Private merr_code2 As Integer
    Private merror_flg As String
    Private mcars As Integer
    Private mtons As Integer
    Private mtotal_rev As Integer
    Private mtc_units As Integer
    Private mserial_no As String


    Public Property Serial_No() As String
        Get
            Return mserial_no
        End Get
        Set(ByVal value As String)
            mserial_no = value
        End Set
    End Property

    Public Property WB_Date() As String
        Get
            Return mwb_date
        End Get
        Set(ByVal value As String)
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

    Public Property Bill_Wght_Tons() As Integer
        Get
            Return mbill_wght_tons
        End Get
        Set(ByVal value As Integer)
            mbill_wght_tons = value
        End Set
    End Property

    Public Property Act_Wght() As Integer
        Get
            Return mact_wght
        End Get
        Set(ByVal value As Integer)
            mact_wght = value
        End Set
    End Property

    Public Property U_Rev() As Integer
        Get
            Return mu_rev
        End Get
        Set(ByVal value As Integer)
            mu_rev = value
        End Set
    End Property

    Public Property Tran_Chrg() As Integer
        Get
            Return mtran_chrg
        End Get
        Set(ByVal value As Integer)
            mtran_chrg = value
        End Set
    End Property

    Public Property Misc_Chrg() As Integer
        Get
            Return mmisc_chrg
        End Get
        Set(ByVal value As Integer)
            mmisc_chrg = value
        End Set
    End Property

    Public Property Intra_State_Code() As Integer
        Get
            Return mintra_state_code
        End Get
        Set(ByVal value As Integer)
            mintra_state_code = value
        End Set
    End Property

    Public Property Transit_Code() As Integer
        Get
            Return mtransit_code
        End Get
        Set(ByVal value As Integer)
            mtransit_code = value
        End Set
    End Property

    Public Property All_Rail_Code() As Integer
        Get
            Return mall_rail_code
        End Get
        Set(ByVal value As Integer)
            mall_rail_code = value
        End Set
    End Property

    Public Property Type_Move() As Integer
        Get
            Return mtype_move
        End Get
        Set(ByVal value As Integer)
            mtype_move = value
        End Set
    End Property

    Public Property Move_Via_Water() As Integer
        Get
            Return mmove_via_water
        End Get
        Set(ByVal value As Integer)
            mmove_via_water = value
        End Set
    End Property

    Public Property Truck_For_Rail() As Integer
        Get
            Return mtruck_for_rail
        End Get
        Set(ByVal value As Integer)
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

    Public Property Rebill() As Integer
        Get
            Return mrebill
        End Get
        Set(ByVal value As Integer)
            mrebill = value
        End Set
    End Property

    Public Property Stratum() As Integer
        Get
            Return mstratum
        End Get
        Set(ByVal value As Integer)
            mstratum = value
        End Set
    End Property

    Public Property Subsample() As Integer
        Get
            Return msubsample
        End Get
        Set(ByVal value As Integer)
            msubsample = value
        End Set
    End Property

    Public Property Report_Period() As Integer
        Get
            Return mreport_period
        End Get
        Set(ByVal value As Integer)
            mreport_period = value
        End Set
    End Property

    Public Property Car_Cap() As Integer
        Get
            Return mcar_cap
        End Get
        Set(ByVal value As Integer)
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

    Public Property Outside_L() As Integer
        Get
            Return moutside_l
        End Get
        Set(ByVal value As Integer)
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

    Public Property Draft_Gear() As Integer
        Get
            Return mdraft_gear
        End Get
        Set(ByVal value As Integer)
            mdraft_gear = value
        End Set
    End Property

    Public Property Art_Units() As Integer
        Get
            Return mart_units
        End Get
        Set(ByVal value As Integer)
            mart_units = value
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

    Public Property STCC() As String
        Get
            Return mstcc
        End Get
        Set(ByVal value As String)
            mstcc = value
        End Set
    End Property

    Public Property JF() As Integer
        Get
            Return mjf
        End Get
        Set(ByVal value As Integer)
            mjf = value
        End Set
    End Property

    Public Property Exp_Factor_Th() As Integer
        Get
            Return mexp_factor_th
        End Get
        Set(ByVal value As Integer)
            mexp_factor_th = value
        End Set
    End Property

    Public Property Error_Flg() As String
        Get
            Return merror_flg
        End Get
        Set(ByVal value As String)
            merror_flg = value
        End Set
    End Property

    Public Property STB_Car_Type() As Integer
        Get
            Return mstb_car_typ
        End Get
        Set(ByVal value As Integer)
            mstb_car_typ = value
        End Set
    End Property

    Public Property Err_Code1() As Integer
        Get
            Return merr_code1
        End Get
        Set(ByVal value As Integer)
            merr_code1 = value
        End Set
    End Property

    Public Property Err_Code2() As Integer
        Get
            Return merr_code2
        End Get
        Set(ByVal value As Integer)
            merr_code2 = value
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

    Public Property Cars() As Integer
        Get
            Return mcars
        End Get
        Set(ByVal value As Integer)
            mcars = value
        End Set
    End Property

    Public Property Tons() As Integer
        Get
            Return mtons
        End Get
        Set(ByVal value As Integer)
            mtons = value
        End Set
    End Property

    Public Property TC_Units() As Integer
        Get
            Return mtc_units
        End Get
        Set(ByVal value As Integer)
            mtc_units = value
        End Set
    End Property

    Public Property Total_Rev() As Integer
        Get
            Return mtotal_rev
        End Get
        Set(ByVal value As Integer)
            mtotal_rev = value
        End Set
    End Property

    Public Property JCT1_ST() As String
        Get
            Return mjct1_st
        End Get
        Set(ByVal value As String)
            mjct1_st = value
        End Set
    End Property

    Public Property JCT2_ST() As String
        Get
            Return mjct2_st
        End Get
        Set(ByVal value As String)
            mjct2_st = value
        End Set
    End Property

    Public Property JCT3_ST() As String
        Get
            Return mjct3_st
        End Get
        Set(ByVal value As String)
            mjct3_st = value
        End Set
    End Property

    Public Property JCT4_ST() As String
        Get
            Return mjct4_st
        End Get
        Set(ByVal value As String)
            mjct4_st = value
        End Set
    End Property

    Public Property JCT5_ST() As String
        Get
            Return mjct5_st
        End Get
        Set(ByVal value As String)
            mjct5_st = value
        End Set
    End Property

    Public Property JCT6_ST() As String
        Get
            Return mjct6_st
        End Get
        Set(ByVal value As String)
            mjct6_st = value
        End Set
    End Property

    Public Property JCT7_ST() As String
        Get
            Return mjct7_st
        End Get
        Set(ByVal value As String)
            mjct7_st = value
        End Set
    End Property

    Public Property JCT8_ST() As String
        Get
            Return mjct8_st
        End Get
        Set(ByVal value As String)
            mjct8_st = value
        End Set
    End Property

    Public Property JCT9_ST() As String
        Get
            Return mjct9_st
        End Get
        Set(ByVal value As String)
            mjct9_st = value
        End Set
    End Property


    Public Property O_BEA() As Integer
        Get
            Return mo_bea
        End Get
        Set(ByVal value As Integer)
            mo_bea = value
        End Set
    End Property

    Public Property T_BEA() As Integer
        Get
            Return mt_bea
        End Get
        Set(ByVal value As Integer)
            mt_bea = value
        End Set
    End Property

    Public Property Exp_Factor() As Integer
        Get
            Return mexp_factor
        End Get
        Set(ByVal value As Integer)
            mexp_factor = value
        End Set
    End Property

    Public Property TOFC_Own_Code() As String
        Get
            Return mtofc_own
        End Get
        Set(ByVal value As String)
            mtofc_own = value
        End Set
    End Property

    Public Property Haz_Bulk() As String
        Get
            Return mhazmat_boxcar
        End Get
        Set(ByVal value As String)
            mhazmat_boxcar = value
        End Set
    End Property

    Public Property O_FT() As Integer
        Get
            Return mo_ft
        End Get
        Set(ByVal value As Integer)
            mo_ft = value
        End Set
    End Property

    Public Property T_FT() As Integer
        Get
            Return mt_ft
        End Get
        Set(ByVal value As Integer)
            mt_ft = value
        End Set
    End Property

End Class
