'''////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>   AAR Index class </summary>
'''
''' <remarks>   Michael Sanders, 10/30/2020. </remarks>
'''////////////////////////////////////////////////////////////////////////////////////////////////////
Public Class Class_URCS_AARIndex

    Private mRegion As Integer
    Private mYear As Integer
    Private mFuel As Single
    Private mMS As Single
    Private mPS As Single
    Private mWage As Single
    Private mMP As Single
    Private mGen As Single

    Public Property Region() As Integer
        Get
            Return mRegion
        End Get
        Set(ByVal mInteger As Integer)
            mRegion = mInteger
        End Set
    End Property

    Public Property Year() As Integer
        Get
            Return mYear
        End Get
        Set(ByVal mInteger As Integer)
            mYear = mInteger
        End Set
    End Property

    Public Property Fuel() As Single
        Get
            Return mFuel
        End Get
        Set(ByVal mSingle As Single)
            mFuel = mSingle
        End Set
    End Property

    Public Property MS() As Single
        Get
            Return mMS
        End Get
        Set(ByVal mSingle As Single)
            mMS = mSingle
        End Set
    End Property

    Public Property PS() As Single
        Get
            Return mPS
        End Get
        Set(ByVal mSingle As Single)
            mPS = mSingle
        End Set
    End Property

    Public Property Wage() As Single
        Get
            Return mWage
        End Get
        Set(ByVal mSingle As Single)
            mWage = mSingle
        End Set
    End Property

    Public Property MP() As Single
        Get
            Return mMP
        End Get
        Set(ByVal mSingle As Single)
            mMP = mSingle
        End Set
    End Property

    Public Property Gen() As Single
        Get
            Return mGen
        End Get
        Set(ByVal mSingle As Single)
            mGen = mSingle
        End Set
    End Property

End Class
