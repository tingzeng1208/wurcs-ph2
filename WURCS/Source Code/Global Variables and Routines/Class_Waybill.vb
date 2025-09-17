'''////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>   Ambigous Waybill class that will be superceded by 913 Waybill class </summary>
'''
''' <remarks>   Michael Sanders, 10/30/2020. </remarks>
'''////////////////////////////////////////////////////////////////////////////////////////////////////
Public Class Class_Waybill

    Private mserial_no As String
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
    Private mint_eq_flg As Byte
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
    Private mstcc As String
    Private morr_alpha As String
    Private mjrr1_alpha As String
    Private mjrr2_alpha As String
    Private mjrr3_alpha As String
    Private mjrr4_alpha As String
    Private mjrr5_alpha As String
    Private mjrr6_alpha As String
    Private mtrr_alpha As String
    Private mjf As Byte
    Private mexp_factor_th As Integer
    Private merror_flg As String
    Private mstb_car_type As Byte
    Private merr_code1 As Byte
    Private merr_code2 As Byte
    Private merr_code3 As Byte
    Private mcar_own As String
    Private mtofc_unit_type As String
    Private mdereg_date As String
    Private mdereg_flg As Byte
    Private mservice_type As Byte
    Private mcars As Integer
    Private mbill_wght_tons As Decimal
    Private mtons As Decimal
    Private mtc_units As Decimal
    Private mtotal_rev As Decimal
    Private morr_rev As Decimal
    Private mjrr1_rev As Decimal
    Private mjrr2_rev As Decimal
    Private mjrr3_rev As Decimal
    Private mjrr4_rev As Decimal
    Private mjrr5_rev As Decimal
    Private mjrr6_rev As Decimal
    Private mtrr_rev As Decimal
    Private morr_dist As Decimal
    Private mjrr1_dist As Decimal
    Private mjrr2_dist As Decimal
    Private mjrr3_dist As Decimal
    Private mjrr4_dist As Decimal
    Private mjrr5_dist As Decimal
    Private mjrr6_dist As Decimal
    Private mtrr_dist As Decimal
    Private mtotal_dist As Decimal
    Private mo_st As String
    Private mjct1_st As String
    Private mjct2_st As String
    Private mjct3_st As String
    Private mjct4_st As String
    Private mjct5_st As String
    Private mjct6_st As String
    Private mjct7_st As String
    Private mt_st As String
    Private mo_bea As Byte
    Private mt_bea As Byte
    Private mo_fips As Long
    Private mt_fips As Long
    Private mo_fa As Byte
    Private mt_fa As Byte
    Private mo_ft As Byte
    Private mt_ft As Byte
    Private mo_smsa As Integer
    Private mt_smsa As Integer
    Private monet As Decimal
    Private mnet1 As Decimal
    Private mnet2 As Decimal
    Private mnet3 As Decimal
    Private mnet4 As Decimal
    Private mnet5 As Decimal
    Private mnet6 As Decimal
    Private mnet7 As Decimal
    Private mtnet As Decimal
    Private mal_flg As Byte
    Private maz_flg As Byte
    Private mar_flg As Byte
    Private mca_flg As Byte
    Private mco_flg As Byte
    Private mct_flg As Byte
    Private mde_flg As Byte
    Private mdc_flg As Byte
    Private mfl_flg As Byte
    Private mga_flg As Byte
    Private mid_flg As Byte
    Private mil_flg As Byte
    Private min_flg As Byte
    Private mia_flg As Byte
    Private mks_flg As Byte
    Private mky_flg As Byte
    Private mla_flg As Byte
    Private mme_flg As Byte
    Private mmd_flg As Byte
    Private mma_flg As Byte
    Private mmi_flg As Byte
    Private mmn_flg As Byte
    Private mms_flg As Byte
    Private mmo_flg As Byte
    Private mmt_flg As Byte
    Private mne_flg As Byte
    Private mnv_flg As Byte
    Private mnh_flg As Byte
    Private mnj_flg As Byte
    Private mnm_flg As Byte
    Private mny_flg As Byte
    Private mnc_flg As Byte
    Private mnd_flg As Byte
    Private moh_flg As Byte
    Private mok_flg As Byte
    Private mor_flg As Byte
    Private mpa_flg As Byte
    Private mri_flg As Byte
    Private msc_flg As Byte
    Private msd_flg As Byte
    Private mtn_flg As Byte
    Private mtx_flg As Byte
    Private mut_flg As Byte
    Private mvt_flg As Byte
    Private mva_flg As Byte
    Private mwa_flg As Byte
    Private mwv_flg As Byte
    Private mwi_flg As Byte
    Private mwy_flg As Byte
    Private mcd_flg As Byte
    Private mmx_flg As Byte
    Private mothr_st_flg As Byte
    Private mint_harm_code As String
    Private mindus_class As String
    Private minter_sic As String
    Private mdom_canada As String
    Private mcs_54 As String
    Private mo_fs_type As String
    Private mt_fs_type As String
    Private mo_fs_ratezip As String
    Private mt_fs_ratezip As String
    Private mo_rate_splc As String
    Private mt_rate_splc As String
    Private mo_swlimit_splc As String
    Private mt_swlimit_splc As String
    Private mo_customs_flg As String
    Private mt_customs_flg As String
    Private mo_grain_flg As String
    Private mt_grain_flg As String
    Private mo_ramp_code As String
    Private mt_ramp_code As String
    Private mo_im_flg As String
    Private mt_im_flg As String
    Private mtransborder_flg As String
    Private morr_Cntry As String
    Private mjrr1_cntry As String
    Private mjrr2_cntry As String
    Private mjrr3_cntry As String
    Private mjrr4_cntry As String
    Private mjrr5_cntry As String
    Private mjrr6_cntry As String
    Private mtrr_cntry As String
    Private mu_fuel_surchrg As Decimal
    Private mo_census_reg As String
    Private mt_census_reg As String
    Private mexp_factor As Decimal
    Private mtotal_vc As Decimal
    Private mrr1_vc As Decimal
    Private mrr2_vc As Decimal
    Private mrr3_vc As Decimal
    Private mrr4_vc As Decimal
    Private mrr5_vc As Decimal
    Private mrr6_vc As Decimal
    Private mrr7_vc As Decimal
    Private mrr8_vc As Decimal
    Private mTracking_No As Long

    Public Property Serial_No() As String
        Get
            Return mserial_no
        End Get
        Set(ByVal value As String)
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

    Public Property Int_Eq_Flg() As Byte
        Get
            Return mint_eq_flg
        End Get
        Set(ByVal value As Byte)
            mint_eq_flg = value
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

    Public Property STCC() As String
        Get
            Return mstcc
        End Get
        Set(ByVal value As String)
            mstcc = value
        End Set
    End Property

    Public Property ORR_Alpha() As String
        Get
            Return morr_alpha
        End Get
        Set(ByVal value As String)
            morr_alpha = value
        End Set
    End Property

    Public Property JRR1_Alpha() As String
        Get
            Return mjrr1_alpha
        End Get
        Set(ByVal value As String)
            mjrr1_alpha = value
        End Set
    End Property

    Public Property JRR2_Alpha() As String
        Get
            Return mjrr2_alpha
        End Get
        Set(ByVal value As String)
            mjrr2_alpha = value
        End Set
    End Property

    Public Property JRR3_Alpha() As String
        Get
            Return mjrr3_alpha
        End Get
        Set(ByVal value As String)
            mjrr3_alpha = value
        End Set
    End Property

    Public Property JRR4_Alpha() As String
        Get
            Return mjrr4_alpha
        End Get
        Set(ByVal value As String)
            mjrr4_alpha = value
        End Set
    End Property

    Public Property JRR5_Alpha() As String
        Get
            Return mjrr5_alpha
        End Get
        Set(ByVal value As String)
            mjrr5_alpha = value
        End Set
    End Property

    Public Property JRR6_Alpha() As String
        Get
            Return mjrr6_alpha
        End Get
        Set(ByVal value As String)
            mjrr6_alpha = value
        End Set
    End Property

    Public Property TRR_Alpha() As String
        Get
            Return mtrr_alpha
        End Get
        Set(ByVal value As String)
            mtrr_alpha = value
        End Set
    End Property

    Public Property JF() As Byte
        Get
            Return mjf
        End Get
        Set(ByVal value As Byte)
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

    Public Property STB_Car_Type() As Byte
        Get
            Return mstb_car_type
        End Get
        Set(ByVal value As Byte)
            mstb_car_type = value
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

    Public Property Dereg_Date() As String
        Get
            Return mdereg_date
        End Get
        Set(ByVal value As String)
            mdereg_date = value
        End Set
    End Property

    Public Property Dereg_Flg() As Byte
        Get
            Return mdereg_flg
        End Get
        Set(ByVal value As Byte)
            mdereg_flg = value
        End Set
    End Property

    Public Property Service_Type() As Byte
        Get
            Return mservice_type
        End Get
        Set(ByVal value As Byte)
            mservice_type = value
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

    Public Property Bill_Wght_Tons() As Decimal
        Get
            Return mbill_wght_tons
        End Get
        Set(ByVal value As Decimal)
            mbill_wght_tons = value
        End Set
    End Property

    Public Property Tons() As Decimal
        Get
            Return mtons
        End Get
        Set(ByVal value As Decimal)
            mtons = value
        End Set
    End Property

    Public Property TC_Units() As Decimal
        Get
            Return mtc_units
        End Get
        Set(ByVal value As Decimal)
            mtc_units = value
        End Set
    End Property

    Public Property Total_Rev() As Decimal
        Get
            Return mtotal_rev
        End Get
        Set(ByVal value As Decimal)
            mtotal_rev = value
        End Set
    End Property

    Public Property ORR_Rev() As Decimal
        Get
            Return morr_rev
        End Get
        Set(ByVal value As Decimal)
            morr_rev = value
        End Set
    End Property

    Public Property JRR1_Rev() As Decimal
        Get
            Return mjrr1_rev
        End Get
        Set(ByVal value As Decimal)
            mjrr1_rev = value
        End Set
    End Property

    Public Property JRR2_Rev() As Decimal
        Get
            Return mjrr2_rev
        End Get
        Set(ByVal value As Decimal)
            mjrr2_rev = value
        End Set
    End Property

    Public Property JRR3_Rev() As Decimal
        Get
            Return mjrr3_rev
        End Get
        Set(ByVal value As Decimal)
            mjrr3_rev = value
        End Set
    End Property

    Public Property JRR4_Rev() As Decimal
        Get
            Return mjrr4_rev
        End Get
        Set(ByVal value As Decimal)
            mjrr4_rev = value
        End Set
    End Property

    Public Property JRR5_Rev() As Decimal
        Get
            Return mjrr5_rev
        End Get
        Set(ByVal value As Decimal)
            mjrr5_rev = value
        End Set
    End Property

    Public Property JRR6_Rev() As Decimal
        Get
            Return mjrr6_rev
        End Get
        Set(ByVal value As Decimal)
            mjrr6_rev = value
        End Set
    End Property

    Public Property TRR_Rev() As Decimal
        Get
            Return mtrr_rev
        End Get
        Set(ByVal value As Decimal)
            mtrr_rev = value
        End Set
    End Property

    Public Property ORR_Dist() As Decimal
        Get
            Return morr_dist
        End Get
        Set(ByVal value As Decimal)
            morr_dist = value
        End Set
    End Property

    Public Property JRR1_Dist() As Decimal
        Get
            Return mjrr1_dist
        End Get
        Set(ByVal value As Decimal)
            mjrr1_dist = value
        End Set
    End Property

    Public Property JRR2_Dist() As Decimal
        Get
            Return mjrr2_dist
        End Get
        Set(ByVal value As Decimal)
            mjrr2_dist = value
        End Set
    End Property

    Public Property JRR3_Dist() As Decimal
        Get
            Return mjrr3_dist
        End Get
        Set(ByVal value As Decimal)
            mjrr3_dist = value
        End Set
    End Property

    Public Property JRR4_Dist() As Decimal
        Get
            Return mjrr4_dist
        End Get
        Set(ByVal value As Decimal)
            mjrr4_dist = value
        End Set
    End Property

    Public Property JRR5_Dist() As Decimal
        Get
            Return mjrr5_dist
        End Get
        Set(ByVal value As Decimal)
            mjrr5_dist = value
        End Set
    End Property

    Public Property JRR6_Dist() As Decimal
        Get
            Return mjrr6_dist
        End Get
        Set(ByVal value As Decimal)
            mjrr6_dist = value
        End Set
    End Property

    Public Property TRR_Dist() As Decimal
        Get
            Return mtrr_dist
        End Get
        Set(ByVal value As Decimal)
            mtrr_dist = value
        End Set
    End Property

    Public Property Total_Dist() As Decimal
        Get
            Return mtotal_dist
        End Get
        Set(ByVal value As Decimal)
            mtotal_dist = value
        End Set
    End Property

    Public Property O_ST() As String
        Get
            Return mo_st
        End Get
        Set(ByVal value As String)
            mo_st = value
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

    Public Property T_ST() As String
        Get
            Return mt_st
        End Get
        Set(ByVal value As String)
            mt_st = value
        End Set
    End Property

    Public Property O_BEA() As Byte
        Get
            Return mo_bea
        End Get
        Set(ByVal value As Byte)
            mo_bea = value
        End Set
    End Property

    Public Property T_BEA() As Byte
        Get
            Return mt_bea
        End Get
        Set(ByVal value As Byte)
            mt_bea = value
        End Set
    End Property

    Public Property O_FIPS() As Long
        Get
            Return mo_fips
        End Get
        Set(ByVal value As Long)
            mo_fips = value
        End Set
    End Property

    Public Property T_FIPS() As Long
        Get
            Return mt_fips
        End Get
        Set(ByVal value As Long)
            mt_fips = value
        End Set
    End Property

    Public Property O_FA() As Byte
        Get
            Return mo_fa
        End Get
        Set(ByVal value As Byte)
            mo_fa = value
        End Set
    End Property

    Public Property T_FA() As Byte
        Get
            Return mt_fa
        End Get
        Set(ByVal value As Byte)
            mt_fa = value
        End Set
    End Property

    Public Property O_FT() As Byte
        Get
            Return mo_ft
        End Get
        Set(ByVal value As Byte)
            mo_ft = value
        End Set
    End Property

    Public Property T_FT() As Byte
        Get
            Return mt_ft
        End Get
        Set(ByVal value As Byte)
            mt_ft = value
        End Set
    End Property

    Public Property O_SMSA() As Integer
        Get
            Return mo_smsa
        End Get
        Set(ByVal value As Integer)
            mo_smsa = value
        End Set
    End Property

    Public Property T_SMSA() As Integer
        Get
            Return mt_smsa
        End Get
        Set(ByVal value As Integer)
            mt_smsa = value
        End Set
    End Property

    Public Property ONET() As Decimal
        Get
            Return monet
        End Get
        Set(ByVal value As Decimal)
            monet = value   'truer words were never typed!
        End Set
    End Property

    Public Property NET1() As Decimal
        Get
            Return mnet1
        End Get
        Set(ByVal value As Decimal)
            mnet1 = value
        End Set
    End Property

    Public Property NET2() As Decimal
        Get
            Return mnet2
        End Get
        Set(ByVal value As Decimal)
            mnet2 = value
        End Set
    End Property

    Public Property NET3() As Decimal
        Get
            Return mnet3
        End Get
        Set(ByVal value As Decimal)
            mnet3 = value
        End Set
    End Property

    Public Property NET4() As Decimal
        Get
            Return mnet4
        End Get
        Set(ByVal value As Decimal)
            mnet4 = value
        End Set
    End Property

    Public Property NET5() As Decimal
        Get
            Return mnet5
        End Get
        Set(ByVal value As Decimal)
            mnet5 = value
        End Set
    End Property

    Public Property NET6() As Decimal
        Get
            Return mnet6
        End Get
        Set(ByVal value As Decimal)
            mnet6 = value
        End Set
    End Property

    Public Property NET7() As Decimal
        Get
            Return mnet7
        End Get
        Set(ByVal value As Decimal)
            mnet7 = value
        End Set
    End Property

    Public Property TNET() As Decimal
        Get
            Return mtnet
        End Get
        Set(ByVal value As Decimal)
            mtnet = value
        End Set
    End Property

    Public Property AL_Flg() As Byte
        Get
            Return mal_flg
        End Get
        Set(ByVal value As Byte)
            mal_flg = value
        End Set
    End Property

    Public Property AZ_Flg() As Byte
        Get
            Return maz_flg
        End Get
        Set(ByVal value As Byte)
            maz_flg = value
        End Set
    End Property

    Public Property AR_Flg() As Byte
        Get
            Return mar_flg
        End Get
        Set(ByVal value As Byte)
            mar_flg = value
        End Set
    End Property

    Public Property CA_Flg() As Byte
        Get
            Return mca_flg
        End Get
        Set(ByVal value As Byte)
            mca_flg = value
        End Set
    End Property

    Public Property CO_Flg() As Byte
        Get
            Return mco_flg
        End Get
        Set(ByVal value As Byte)
            mco_flg = value
        End Set
    End Property

    Public Property CT_Flg() As Byte
        Get
            Return mct_flg
        End Get
        Set(ByVal value As Byte)
            mct_flg = value
        End Set
    End Property

    Public Property DE_Flg() As Byte
        Get
            Return mde_flg
        End Get
        Set(ByVal value As Byte)
            mde_flg = value
        End Set
    End Property

    Public Property DC_Flg() As Byte
        Get
            Return mdc_flg
        End Get
        Set(ByVal value As Byte)
            mdc_flg = value
        End Set
    End Property

    Public Property FL_Flg() As Byte
        Get
            Return mfl_flg
        End Get
        Set(ByVal value As Byte)
            mfl_flg = value
        End Set
    End Property

    Public Property GA_Flg() As Byte
        Get
            Return mga_flg
        End Get
        Set(ByVal value As Byte)
            mga_flg = value
        End Set
    End Property

    Public Property ID_Flg() As Byte
        Get
            Return mid_flg
        End Get
        Set(ByVal value As Byte)
            mid_flg = value
        End Set
    End Property

    Public Property IL_Flg() As Byte
        Get
            Return mil_flg
        End Get
        Set(ByVal value As Byte)
            mil_flg = value
        End Set
    End Property

    Public Property IN_Flg() As Byte
        Get
            Return min_flg
        End Get
        Set(ByVal value As Byte)
            min_flg = value
        End Set
    End Property

    Public Property IA_Flg() As Byte
        Get
            Return mia_flg
        End Get
        Set(ByVal value As Byte)
            mia_flg = value
        End Set
    End Property

    Public Property KS_Flg() As Byte
        Get
            Return mks_flg
        End Get
        Set(ByVal value As Byte)
            mks_flg = value
        End Set
    End Property

    Public Property KY_Flg() As Byte
        Get
            Return mky_flg
        End Get
        Set(ByVal value As Byte)
            mky_flg = value
        End Set
    End Property

    Public Property LA_Flg() As Byte
        Get
            Return mla_flg
        End Get
        Set(ByVal value As Byte)
            mla_flg = value
        End Set
    End Property

    Public Property ME_Flg() As Byte
        Get
            Return mme_flg
        End Get
        Set(ByVal value As Byte)
            mme_flg = value
        End Set
    End Property

    Public Property MD_Flg() As Byte
        Get
            Return mmd_flg
        End Get
        Set(ByVal value As Byte)
            mmd_flg = value
        End Set
    End Property

    Public Property MA_Flg() As Byte
        Get
            Return mma_flg
        End Get
        Set(ByVal value As Byte)
            mma_flg = value
        End Set
    End Property

    Public Property MI_Flg() As Byte
        Get
            Return mmi_flg
        End Get
        Set(ByVal value As Byte)
            mmi_flg = value
        End Set
    End Property

    Public Property MN_Flg() As Byte
        Get
            Return mmn_flg
        End Get
        Set(ByVal value As Byte)
            mmn_flg = value
        End Set
    End Property

    Public Property MS_Flg() As Byte
        Get
            Return mms_flg
        End Get
        Set(ByVal value As Byte)
            mms_flg = value
        End Set
    End Property

    Public Property MO_Flg() As Byte
        Get
            Return mmo_flg
        End Get
        Set(ByVal value As Byte)
            mmo_flg = value
        End Set
    End Property

    Public Property MT_Flg() As Byte
        Get
            Return mmt_flg
        End Get
        Set(ByVal value As Byte)
            mmt_flg = value
        End Set
    End Property

    Public Property NE_Flg() As Byte
        Get
            Return mne_flg
        End Get
        Set(ByVal value As Byte)
            mne_flg = value
        End Set
    End Property

    Public Property NV_Flg() As Byte
        Get
            Return mnv_flg
        End Get
        Set(ByVal value As Byte)
            mnv_flg = value
        End Set
    End Property

    Public Property NH_Flg() As Byte
        Get
            Return mnh_flg
        End Get
        Set(ByVal value As Byte)
            mnh_flg = value
        End Set
    End Property

    Public Property NJ_Flg() As Byte
        Get
            Return mnj_flg
        End Get
        Set(ByVal value As Byte)
            mnj_flg = value
        End Set
    End Property

    Public Property NM_Flg() As Byte
        Get
            Return mnm_flg
        End Get
        Set(ByVal value As Byte)
            mnm_flg = value
        End Set
    End Property

    Public Property NY_Flg() As Byte
        Get
            Return mny_flg
        End Get
        Set(ByVal value As Byte)
            mny_flg = value
        End Set
    End Property

    Public Property NC_Flg() As Byte
        Get
            Return mnc_flg
        End Get
        Set(ByVal value As Byte)
            mnc_flg = value
        End Set
    End Property

    Public Property ND_Flg() As Byte
        Get
            Return mnd_flg
        End Get
        Set(ByVal value As Byte)
            mnd_flg = value
        End Set
    End Property

    Public Property OH_Flg() As Byte
        Get
            Return moh_flg
        End Get
        Set(ByVal value As Byte)
            moh_flg = value
        End Set
    End Property

    Public Property OK_Flg() As Byte
        Get
            Return mok_flg
        End Get
        Set(ByVal value As Byte)
            mok_flg = value
        End Set
    End Property

    Public Property OR_Flg() As Byte
        Get
            Return mor_flg
        End Get
        Set(ByVal value As Byte)
            mor_flg = value
        End Set
    End Property

    Public Property PA_Flg() As Byte
        Get
            Return mpa_flg
        End Get
        Set(ByVal value As Byte)
            mpa_flg = value
        End Set
    End Property

    Public Property RI_Flg() As Byte
        Get
            Return mri_flg
        End Get
        Set(ByVal value As Byte)
            mri_flg = value
        End Set
    End Property

    Public Property SC_Flg() As Byte
        Get
            Return msc_flg
        End Get
        Set(ByVal value As Byte)
            msc_flg = value
        End Set
    End Property

    Public Property SD_Flg() As Byte
        Get
            Return msd_flg
        End Get
        Set(ByVal value As Byte)
            msd_flg = value
        End Set
    End Property

    Public Property TN_Flg() As Byte
        Get
            Return mtn_flg
        End Get
        Set(ByVal value As Byte)
            mtn_flg = value
        End Set
    End Property

    Public Property TX_Flg() As Byte
        Get
            Return mtx_flg
        End Get
        Set(ByVal value As Byte)
            mtx_flg = value
        End Set
    End Property

    Public Property UT_Flg() As Byte
        Get
            Return mut_flg
        End Get
        Set(ByVal value As Byte)
            mut_flg = value
        End Set
    End Property

    Public Property VT_Flg() As Byte
        Get
            Return mvt_flg
        End Get
        Set(ByVal value As Byte)
            mvt_flg = value
        End Set
    End Property

    Public Property VA_Flg() As Byte
        Get
            Return mva_flg
        End Get
        Set(ByVal value As Byte)
            mva_flg = value
        End Set
    End Property

    Public Property WA_Flg() As Byte
        Get
            Return mwa_flg
        End Get
        Set(ByVal value As Byte)
            mwa_flg = value
        End Set
    End Property

    Public Property WV_Flg() As Byte
        Get
            Return mwv_flg
        End Get
        Set(ByVal value As Byte)
            mwv_flg = value
        End Set
    End Property

    Public Property WI_Flg() As Byte
        Get
            Return mwi_flg
        End Get
        Set(ByVal value As Byte)
            mwi_flg = value
        End Set
    End Property

    Public Property WY_Flg() As Byte
        Get
            Return mwy_flg
        End Get
        Set(ByVal value As Byte)
            mwy_flg = value
        End Set
    End Property

    Public Property CD_Flg() As Byte
        Get
            Return mcd_flg
        End Get
        Set(ByVal value As Byte)
            mcd_flg = value
        End Set
    End Property

    Public Property MX_Flg() As Byte
        Get
            Return mmx_flg
        End Get
        Set(ByVal value As Byte)
            mmx_flg = value
        End Set
    End Property

    Public Property Othr_St_Flg() As Byte
        Get
            Return mothr_st_flg
        End Get
        Set(ByVal value As Byte)
            mothr_st_flg = value
        End Set
    End Property

    Public Property Int_Harm_Code() As String
        Get
            Return mint_harm_code
        End Get
        Set(ByVal value As String)
            mint_harm_code = value
        End Set
    End Property

    Public Property Indus_Class() As String
        Get
            Return mindus_class
        End Get
        Set(ByVal value As String)
            mindus_class = value
        End Set
    End Property

    Public Property Inter_Sic() As String
        Get
            Return minter_sic
        End Get
        Set(ByVal value As String)
            minter_sic = value
        End Set
    End Property

    Public Property Dom_Canada() As String
        Get
            Return mdom_canada
        End Get
        Set(ByVal value As String)
            mdom_canada = value
        End Set
    End Property

    Public Property CS_54() As String
        Get
            Return mcs_54
        End Get
        Set(ByVal value As String)
            mcs_54 = value
        End Set
    End Property

    Public Property O_FS_Type() As String
        Get
            Return mo_fs_type
        End Get
        Set(ByVal value As String)
            mo_fs_type = value
        End Set
    End Property

    Public Property T_FS_Type() As String
        Get
            Return mt_fs_type
        End Get
        Set(ByVal value As String)
            mt_fs_type = value
        End Set
    End Property

    Public Property O_FS_RateZip() As String
        Get
            Return mo_fs_ratezip
        End Get
        Set(ByVal value As String)
            mo_fs_ratezip = value
        End Set
    End Property

    Public Property T_FS_RateZip() As String
        Get
            Return mt_fs_ratezip
        End Get
        Set(ByVal value As String)
            mt_fs_ratezip = value
        End Set
    End Property

    Public Property O_Rate_SPLC() As String
        Get
            Return mo_rate_splc
        End Get
        Set(ByVal value As String)
            mo_rate_splc = value
        End Set
    End Property

    Public Property T_Rate_SPLC() As String
        Get
            Return mt_rate_splc
        End Get
        Set(ByVal value As String)
            mt_rate_splc = value
        End Set
    End Property

    Public Property O_SWLimit_SPLC() As String
        Get
            Return mo_swlimit_splc
        End Get
        Set(ByVal value As String)
            mo_swlimit_splc = value
        End Set
    End Property

    Public Property T_SWLimit_SPLC() As String
        Get
            Return mt_swlimit_splc
        End Get
        Set(ByVal value As String)
            mt_swlimit_splc = value
        End Set
    End Property

    Public Property O_Customs_Flg() As String
        Get
            Return mo_customs_flg
        End Get
        Set(ByVal value As String)
            mo_customs_flg = value
        End Set
    End Property

    Public Property T_Customs_Flg() As String
        Get
            Return mt_customs_flg
        End Get
        Set(ByVal value As String)
            mt_customs_flg = value
        End Set
    End Property

    Public Property O_Grain_Flg() As String
        Get
            Return mo_grain_flg
        End Get
        Set(ByVal value As String)
            mo_grain_flg = value
        End Set
    End Property

    Public Property T_Grain_Flg() As String
        Get
            Return mt_grain_flg
        End Get
        Set(ByVal value As String)
            mt_grain_flg = value
        End Set
    End Property

    Public Property O_Ramp_Code() As String
        Get
            Return mo_ramp_code
        End Get
        Set(ByVal value As String)
            mo_ramp_code = value
        End Set
    End Property

    Public Property T_Ramp_Code() As String
        Get
            Return mt_ramp_code
        End Get
        Set(ByVal value As String)
            mt_ramp_code = value
        End Set
    End Property

    Public Property O_IM_Flg() As String
        Get
            Return mo_im_flg
        End Get
        Set(ByVal value As String)
            mo_im_flg = value
        End Set
    End Property

    Public Property T_IM_Flg() As String
        Get
            Return mt_im_flg
        End Get
        Set(ByVal value As String)
            mt_im_flg = value
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

    Public Property ORR_Cntry() As String
        Get
            Return morr_Cntry
        End Get
        Set(ByVal value As String)
            morr_Cntry = value
        End Set
    End Property

    Public Property JRR1_Cntry() As String
        Get
            Return mjrr1_cntry
        End Get
        Set(ByVal value As String)
            mjrr1_cntry = value
        End Set
    End Property

    Public Property JRR2_Cntry() As String
        Get
            Return mjrr2_cntry
        End Get
        Set(ByVal value As String)
            mjrr2_cntry = value
        End Set
    End Property

    Public Property JRR3_Cntry() As String
        Get
            Return mjrr3_cntry
        End Get
        Set(ByVal value As String)
            mjrr3_cntry = value
        End Set
    End Property

    Public Property JRR4_Cntry() As String
        Get
            Return mjrr4_cntry
        End Get
        Set(ByVal value As String)
            mjrr4_cntry = value
        End Set
    End Property

    Public Property JRR5_Cntry() As String
        Get
            Return mjrr5_cntry
        End Get
        Set(ByVal value As String)
            mjrr5_cntry = value
        End Set
    End Property

    Public Property JRR6_Cntry() As String
        Get
            Return mjrr6_cntry
        End Get
        Set(ByVal value As String)
            mjrr6_cntry = value
        End Set
    End Property

    Public Property TRR_Cntry() As String
        Get
            Return mtrr_cntry
        End Get
        Set(ByVal value As String)
            mtrr_cntry = value
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

    Public Property O_Census_Reg() As String
        Get
            Return mo_census_reg
        End Get
        Set(ByVal value As String)
            mo_census_reg = value
        End Set
    End Property

    Public Property T_Census_Reg() As String
        Get
            Return mt_census_reg
        End Get
        Set(ByVal value As String)
            mt_census_reg = value
        End Set
    End Property

    Public Property Exp_Factor() As Decimal
        Get
            Return mexp_factor
        End Get
        Set(ByVal value As Decimal)
            mexp_factor = value
        End Set
    End Property

    Public Property Total_VC() As Decimal
        Get
            Return mtotal_vc
        End Get
        Set(ByVal value As Decimal)
            mtotal_vc = value
        End Set
    End Property

    Public Property RR1_VC() As Decimal
        Get
            Return mrr1_vc
        End Get
        Set(ByVal value As Decimal)
            mrr1_vc = value
        End Set
    End Property

    Public Property RR2_VC() As Decimal
        Get
            Return mrr2_vc
        End Get
        Set(ByVal value As Decimal)
            mrr2_vc = value
        End Set
    End Property

    Public Property RR3_VC() As Decimal
        Get
            Return mrr3_vc
        End Get
        Set(ByVal value As Decimal)
            mrr3_vc = value
        End Set
    End Property

    Public Property RR4_VC() As Decimal
        Get
            Return mrr4_vc
        End Get
        Set(ByVal value As Decimal)
            mrr4_vc = value
        End Set
    End Property

    Public Property RR5_VC() As Decimal
        Get
            Return mrr5_vc
        End Get
        Set(ByVal value As Decimal)
            mrr5_vc = value
        End Set
    End Property

    Public Property RR6_VC() As Decimal
        Get
            Return mrr6_vc
        End Get
        Set(ByVal value As Decimal)
            mrr6_vc = value
        End Set
    End Property

    Public Property RR7_VC() As Decimal
        Get
            Return mrr7_vc
        End Get
        Set(ByVal value As Decimal)
            mrr7_vc = value
        End Set
    End Property

    Public Property RR8_VC() As Decimal
        Get
            Return mrr8_vc
        End Get
        Set(ByVal value As Decimal)
            mrr8_vc = value
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
