
Namespace VBFastTag.PosTag
    ''' <summary>
    ''' Represent POS result for a word
    ''' </summary>
    Public Class FastTagResult
        ''' <summary>
        ''' The word used for tagging
        ''' </summary>

        Public Property Word() As String
            Get
                Return m_Word
            End Get
            Private Set(value As String)
                m_Word = Value
            End Set
        End Property
        Private m_Word As String

        ''' <summary>
        ''' The assigned tag
        ''' </summary>

        Public Property PosTag() As String
            Get
                Return m_PosTag
            End Get
            Private Set(value As String)
                m_PosTag = Value
            End Set
        End Property
        Private m_PosTag As String

        ''' <summary>
        ''' CTOR
        ''' </summary>
        ''' <param name="word__1"></param>
        ''' <param name="pTag"></param>

        Public Sub New(word__1 As String, pTag As String)
            Word = word__1
            PosTag = pTag
        End Sub
    End Class
End Namespace

