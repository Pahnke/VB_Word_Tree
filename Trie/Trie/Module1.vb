Module Module1
    Public Const NO_LETTERS As Integer = 26
    Public Const FIRST_LETTER As Integer = Asc("A")

    Class Tree_Node
        Private children(NO_LETTERS) As Tree_Node
        Private word As Boolean

        Public Sub New()
            word = False
        End Sub

        Public Function Remove(ByVal value As String) As Boolean
            Dim final_node As Tree_Node = Find_Node(value)
            If Not (final_node.word) Then
                Return False
            End If
            final_node.word = False
            Return True
        End Function

        Public Function Check_In_Tree(ByVal value As String) As Boolean
            Return Find_Node(value).word
        End Function

        Private Function Find_Node(ByVal value As String) As Tree_Node
            If value = "" Then
                Return Me
            End If
            Dim position As Integer = Asc(UCase(value(0))) - FIRST_LETTER
            If position < 0 Or position >= NO_LETTERS Then
                Return Nothing
            End If
            Create_Child(position)
            Return children(position).Find_Node(value.Substring(1))
        End Function

        Public Function Add(ByVal value As String) As Boolean
            Dim final_node As Tree_Node = Find_Node(value)
            If final_node.word Then
                Return False
            End If
            final_node.word = True
            Return True
        End Function

        Private Sub Create_Child(ByVal position As Integer)
            If children(position) Is Nothing Then
                children(position) = New Tree_Node
            End If
        End Sub

    End Class


    Sub Main()
        Dim root As Tree_Node = New Tree_Node


    End Sub

End Module

