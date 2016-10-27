
' VBFastTag is a VB.NET port of Mark Watson's FastTag Part of Speech Tagger which was itself 
'based on Eric Brill's trained rule set and English lexicon.
'Licensed under LGPL3 or Apache 2 licenses
' For an alternative non-GPL license: contact the author
' THIS SOFTWARE COMES WITH NO WARRANTY

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace VBFastTag.PosTag
    ''' <summary>
    ''' VB.NET port of the mark-watson FastTag_v2
    ''' </summary>
    Public Class FastTag
        ''' <summary>
        ''' Internal word lexicon where the word/pos tags are stored
        ''' </summary>
        Private ReadOnly _lexicon As New Dictionary(Of String, [String]())()

        ''' <summary>
        ''' CTOR
        ''' </summary>
        ''' <param name="learningData"></param>
        Public Sub New(learningData As String)
            Using sr = New StringReader(learningData)
                Dim line As String
                line = ""
                While (InlineAssignHelper(line, sr.ReadLine())) IsNot Nothing
                    ParseLine(line)
                End While
            End Using
        End Sub

        ''' <summary>
        ''' Checks if the provided word exist in the imported lexicon
        ''' </summary>
        ''' <param name="word"></param>
        ''' <returns></returns>
        Public Function WordInLexicon(word As String) As Boolean
            Return _lexicon.ContainsKey(word) OrElse _lexicon.ContainsKey(word.ToLower())
        End Function

        ''' <summary>
        ''' Assigns parts of speech to each word
        ''' </summary>
        ''' <param name="words"></param>
        ''' <returns></returns>
        Public Function Tag(words As IList(Of String)) As IList(Of FastTagResult)
            If words Is Nothing OrElse words.Count = 0 Then
                Return New List(Of FastTagResult)()
            End If

            Dim result = New List(Of FastTagResult)()
            Dim pTags = GetPosTagsFor(words)
            ' Apply transformational rules
            For i As Integer = 0 To words.Count - 1
                Dim word As String = words(i)
                Dim pTag As String = pTags(i)
                '  rule 1: DT, {VBD | VBP} --> DT, NN
                If i > 0 AndAlso String.Equals(pTags(i - 1), "DT") Then
                    If String.Equals(pTag, "VBD") OrElse String.Equals(pTag, "VBP") OrElse String.Equals(pTag, "VB") Then
                        pTag = "NN"
                    End If
                End If
                ' rule 2: convert a noun to a number (CD) if "." appears in the word
                If pTag.StartsWith("N") Then
                    Dim s As [Single]
                    If word.IndexOf(".", StringComparison.CurrentCultureIgnoreCase) > -1 OrElse [Single].TryParse(word, s) Then
                        pTag = "CD"
                    End If
                End If
                ' rule 3: convert a noun to a past participle if words.get(i) ends with "ed"
                If pTag.StartsWith("N") AndAlso word.EndsWith("ed") Then
                    pTag = "VBN"
                End If
                ' rule 4: convert any type to adverb if it ends in "ly";
                If word.EndsWith("ly") Then
                    pTag = "RB"
                End If
                ' rule 5: convert a common noun (NN or NNS) to a adjective if it ends with "al"
                If pTag.StartsWith("NN") AndAlso word.EndsWith("al") Then
                    pTag = "JJ"
                End If
                ' rule 6: convert a noun to a verb if the preceeding work is "would"
                If i > 0 AndAlso pTag.StartsWith("NN") AndAlso String.Equals(words(i - 1), "would") Then
                    pTag = "VB"
                End If
                ' rule 7: if a word has been categorized as a common noun and it ends with "s",
                '         then set its type to plural common noun (NNS)
                If String.Equals(pTag, "NN") AndAlso word.EndsWith("s") Then
                    pTag = "NNS"
                End If
                ' rule 8: convert a common noun to a present participle verb (i.e., a gerand)
                If pTag.StartsWith("NN") AndAlso word.EndsWith("ing") Then
                    pTag = "VBG"
                End If

                result.Add(New FastTagResult(word, pTag))
            Next
            Return result
        End Function

        ''' <summary>
        ''' Assigns parts of speech to a sentence
        ''' </summary>
        ''' <param name="sentence"></param>
        ''' <returns></returns>
        Public Function Tag(sentence As String) As IList(Of FastTagResult)
            If String.IsNullOrEmpty(sentence) Then
                Return New List(Of FastTagResult)()
            End If
            Dim sentenceWords = sentence.Split(" "c)
            Return Tag(sentenceWords)
        End Function

        ''' <summary>
        ''' Retrieve the pos tags from the lexicon for the provided word list
        ''' </summary>
        ''' <param name="words"></param>
        ''' <returns></returns>
        Private Function GetPosTagsFor(words As IList(Of String)) As IList(Of String)
            Dim ret As IList(Of String) = New List(Of String)(words.Count)
            Dim i As Integer = 0, size As Integer = words.Count
            While i < size
                Dim word = RemoveSpecialCharacters(words(i))
                If String.IsNullOrEmpty(word) Then
                    ret.Add("")
                    Continue While
                End If

                Dim ss As String() = {}


                _lexicon.TryGetValue(word, ss)
                ' 1/22/2002 mod (from Lisp code): if not in hash, try lower case:
                If ss Is Nothing Then
                    _lexicon.TryGetValue(word.ToLower(), ss)
                End If
                If ss Is Nothing AndAlso word.Length = 1 Then
                    ret.Add(word & Convert.ToString("^"))
                ElseIf ss Is Nothing Then
                    ret.Add("NN")
                Else
                    ret.Add(ss(0))
                End If
                i += 1
            End While

            Return ret
        End Function

        ''' <summary>
        ''' Clears special chars from start and end of the word
        ''' </summary>
        ''' <param name="str"></param>
        ''' <returns></returns>
        Private Function RemoveSpecialCharacters(str As String) As String
            If str.Length = 1 Then
                Return str
            End If
            Dim rpl = Regex.Replace(str, "^[^A-Za-z0-9]+|[^A-Za-z0-9]+$", String.Empty)
            Return rpl
        End Function

        ''' <summary>
        ''' Parse a line into word and part of speech tags
        ''' </summary>
        ''' <param name="line"></param>
        Private Sub ParseLine(line As String)
            Dim ss = line.Split(" "c)
            Dim word = ss(0)
            _lexicon(word) = ss.Skip(1).ToArray()
        End Sub
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class
End Namespace
