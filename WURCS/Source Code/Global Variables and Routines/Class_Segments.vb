'''////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>   Waybill Segments class </summary>
'''
''' <remarks>   Michael Sanders, 10/30/2020. </remarks>
'''////////////////////////////////////////////////////////////////////////////////////////////////////
Public Class Class_Segments

    Private mSerial_No As String
    Private mSegment_No As Decimal
    Private mTotal_Segments As Decimal
    Private mRR_Num As Integer
    Private mRR_Alpha As String
    Private mRR_Dist As Single
    Private mRR_Cntry As String
    Private mRR_Rev As Decimal
    Private mRR_VC As Decimal
    Private mSeg_Type As String
    Private mFrom_Node As Single
    Private mTo_Node As Single
    Private mFrom_Loc As String
    Private mFrom_St As String
    Private mTo_Loc As String
    Private mTo_St As String

    Public Property Serial_No() As Decimal
        Get
            Return mSerial_No
        End Get
        Set(ByVal value As Decimal)
            mSerial_No = value
        End Set
    End Property

    Public Property Segment_No() As Decimal
        Get
            Return mSegment_No
        End Get
        Set(ByVal value As Decimal)
            mSegment_No = value
        End Set
    End Property

    Public Property Total_Segments() As Decimal
        Get
            Return mTotal_Segments
        End Get
        Set(ByVal value As Decimal)
            mTotal_Segments = value
        End Set
    End Property

    Public Property RR_Num() As Integer
        Get
            Return mRR_Num
        End Get
        Set(ByVal value As Integer)
            mRR_Num = value
        End Set
    End Property

    Public Property RR_Alpha() As String
        Get
            Return mRR_Alpha
        End Get
        Set(ByVal value As String)
            mRR_Alpha = value
        End Set
    End Property

    Public Property RR_Dist() As Single
        Get
            Return mRR_Dist
        End Get
        Set(ByVal value As Single)
            mRR_Dist = value
        End Set
    End Property

    Public Property RR_Cntry() As String
        Get
            Return mRR_Cntry
        End Get
        Set(ByVal value As String)
            mRR_Cntry = value
        End Set
    End Property

    Public Property RR_Rev() As Decimal
        Get
            Return mRR_Rev
        End Get
        Set(ByVal value As Decimal)
            mRR_Rev = value
        End Set
    End Property

    Public Property RR_VC() As Decimal
        Get
            Return mRR_VC
        End Get
        Set(ByVal value As Decimal)
            mRR_VC = value
        End Set
    End Property

    Public Property Seg_Type() As String
        Get
            Return mSeg_Type
        End Get
        Set(ByVal value As String)
            mSeg_Type = value
        End Set
    End Property

    Public Property From_Node() As Decimal
        Get
            Return mFrom_Node
        End Get
        Set(ByVal value As Decimal)
            mFrom_Node = value
        End Set
    End Property

    Public Property To_Node() As Decimal
        Get
            Return mTo_Node
        End Get
        Set(ByVal value As Decimal)
            mTo_Node = value
        End Set
    End Property

    Public Property From_Loc() As String
        Get
            Return mFrom_Loc
        End Get
        Set(ByVal value As String)
            mFrom_Loc = value
        End Set
    End Property

    Public Property From_St() As String
        Get
            Return mFrom_St
        End Get
        Set(ByVal value As String)
            mFrom_St = value
        End Set
    End Property

    Public Property To_Loc() As String
        Get
            Return mTo_Loc
        End Get
        Set(ByVal value As String)
            mTo_Loc = value
        End Set
    End Property

    Public Property To_St() As String
        Get
            Return mTo_St
        End Get
        Set(ByVal value As String)
            mTo_St = value
        End Set
    End Property

End Class
