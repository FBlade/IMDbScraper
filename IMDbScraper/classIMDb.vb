Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Net
Imports System.IO

Public Class IMDb

#Region "> ======== ASSEMBLY INFO ========"
    Public Shared ReadOnly Property libraryName() As String
        Get
            Return Title & " " & Version.ToString(".00")
        End Get
    End Property

    Public Shared ReadOnly Property Title() As String
        Get
            Return CType(Reflection.AssemblyTitleAttribute.GetCustomAttribute( _
            Reflection.Assembly.GetExecutingAssembly, _
            GetType(Reflection.AssemblyTitleAttribute)), Reflection.AssemblyTitleAttribute).Title
        End Get
    End Property

    Public Shared ReadOnly Property Version() As Decimal
        Get
            With Reflection.Assembly.GetExecutingAssembly.GetName.Version
                Return CDec(.Major & "." & .Minor)
            End With
        End Get
    End Property
#End Region

#Region "> ======== CONSTRUCTOR ========"
    Const userAgent As String = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:26.0) Gecko/20100101 Firefox/26.0"
    Const accePt As String = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    Const acceptCharset As String = "ISO-8859-1,utf-8;q=0.7,*;q=0.7"
    Const acceptLanguage As String = "en-us,en;q=0.5"
    Private iSource As String

    Sub New(ByVal iSource As String)
        Me.iSource = Regex.Replace(iSource, "&nbsp;", " ")
    End Sub

    Sub New(ByVal movNo As String, Optional ByVal timeOut As Integer = 5000)
        Dim requestScrape As HttpWebRequest = WebRequest.Create("http://www.imdb.com/title/tt" & movNo & "/combined")
        Dim responseScrape As HttpWebResponse
        Dim sReader As StreamReader
        requestScrape.KeepAlive = False
        requestScrape.Timeout = timeOut
        requestScrape.ReadWriteTimeout = timeOut * 1.5
        requestScrape.Accept = accePt
        requestScrape.Headers.Add("Accept-Charset", acceptCharset)
        requestScrape.Headers.Add("Accept-Language", acceptLanguage)
        requestScrape.UserAgent = userAgent
        requestScrape.AutomaticDecompression = DecompressionMethods.GZip
        responseScrape = requestScrape.GetResponse
        sReader = New StreamReader(responseScrape.GetResponseStream)
        iSource = Regex.Replace(sReader.ReadToEnd, "&nbsp;", " ")
        sReader.Close()
        responseScrape.Close()
    End Sub
#End Region

#Region "> ======== SHARED FUNCTION ========"
    Public Shared Function linkValidator(ByVal url As String) As Boolean
        Return Regex.IsMatch(url.Trim, "\A(http://\w+.imdb.\w+/\w+/tt\d{7}/?|tt\d{7}\b)")
    End Function

    Public Shared Function getMovNo(ByVal url As String) As String
        Return Regex.Match(url.Trim, "\A(http://\w+.imdb.\w+/\w+/tt|tt)(?<mov>\d{7})").Groups("mov").Value
    End Function
#End Region

#Region "> ======== GET ========"
    Public ReadOnly Property getSource() As String
        Get
            Return iSource
        End Get
    End Property

    Public Function getTitle() As String
        ''Title
        'tt1375666 NORMAL <div id="tn15title">' \n '<h1>Inception <span>(<a href="/year/2010/">2010</a>) <span class="pro-link"><a href="">More at <strong>IMDbPro</strong></a>&nbsp;&raquo;</span>'
        '                 <span class="title-extra"></span></span></h1>' \n '</div>
        'tt0655454 EM     <div id="tn15title">' \n '<h1>&#x22;Mystery Science Theater 3000&#x22; <span><em>The Beast of Yucca Flats</em> (1995)</span>'
        '                 <span class="title-extra"></span></h1>' \n '</div>
        'tt1220719 ALT    <div id="tn15title">' \n '<h1>Ip Man <span>(<a href="/year/2008/">2008</a>) <span class="pro-link"><a href="">More at <strong>IMDbPro</strong></a>&nbsp;&raquo;</span>'
        '                 <span class="title-extra">Yip Man <i>(original title)</i></span></span></h1>' \n '</div>

        'revision 1.02 : Match = Regex.Match(httpImdbSource, "<div id=""tn15title"">\s*<\w+>(?<title>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)<\w+>\s*(<em>(?<em>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)</em>\s*)?(\((<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>)?(?<year>\d{4}))?", RegexOptions.IgnoreCase)
        'revision 1.31 : Match = Regex.Match(iSource, "h1>\s*(?<title>((?!<span|h1).)+)<span>(\s*<em>(?<em>((?!</em|h1).)+)?</em>)?", RegexOptions.IgnoreCase)
        'revision 1.32 : SEPERATE INTO getMovTitle & getAltTitle
        Dim strTitle As String = String.Empty
        If getAltTitle() = String.Empty Then : strTitle = getMovTitle()
        Else : strTitle = getAltTitle() : End If
        Return charRef.cleanRef(strTitle)

        'revision 1.31
        'Dim Match As Match = Regex.Match(iSource, "h1>\s*(?<title>((?!<span|h1).)+)<span>(\s*<em>(?<em>((?!</em|h1).)+)?</em>)?(\s*(?<extra>((?!</i|</h1|<div).)+)?<)?", RegexOptions.IgnoreCase)
        'Dim strTitle As String = String.Empty
        'If Match.Groups("title").Success Then
        '    strTitle = Match.Groups("title").Value.TrimEnd
        '    If Match.Groups("em").Success Then
        '        strTitle &= " " & Match.Groups("em").Value.TrimEnd
        '    End If
        '    If Match.Groups("extra").Success Then
        '        Dim altTitle As String = getAltTitle(Match.Groups("extra").Value)
        '        If altTitle <> String.Empty Then : strTitle = altTitle : End If
        '    End If
        'End If
        'Return charRef.cleanRef(strTitle)

        'revision 1.02 : ''Original Title
        ''<div id="tn15title">' vbNewLine '<h1>Triple Tap <span>(<a href="/year/2010/">2010</a>)
        '' .... '<span class="title-extra">Cheung wong chi wong <i>(original title)</i></span> 
        'Match = Regex.Match(httpImdbSource, "<span class=""title\-extra"">(?<orititle>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)", RegexOptions.IgnoreCase)
        'If Match.Groups("orititle").Success Then
        '    info.AlsoKnownAs = """" & info.EnTitle & """ - International (english title)" & vbNewLine
        '    info.EnTitle = charRef.addRef(Match.Groups("orititle").Captures.Item(0).Value).Trim
        'End If
    End Function

    Private Function getMovTitle() As String
        Dim Match As Match = Regex.Match(iSource, "h1>\s*(?<title>((?!<span|h1).)+)<span>(\s*<em>(?<em>((?!</em|h1).)+)?</em>)?", RegexOptions.IgnoreCase)
        Dim strTitle As String = String.Empty
        If Match.Groups("title").Success Then
            strTitle = Match.Groups("title").Value.TrimEnd
            If Match.Groups("em").Success Then
                strTitle &= " - " & Match.Groups("em").Value
            End If
        End If
        Return Regex.Replace(strTitle, "<((?<!>).)+", String.Empty).TrimEnd
    End Function

    Private Function getAltTitle() As String
        Dim Match As Match = Regex.Match(iSource, "h1>\s*(?<inner>((?!</h1|<div).)+)?", RegexOptions.IgnoreCase)
        If Match.Groups("inner").Success Then
            Dim alt As Match = Regex.Match(Match.Groups("inner").Value, "<span class=""title\-extra"">\s*(?!</span)(?<extra>((?!</span|</h1|<div).)*)?", RegexOptions.IgnoreCase)
            If alt.Groups("extra").Success Then
                Return Regex.Replace(alt.Groups("extra").Value, "[<\(]((?<!>|\)).)+", String.Empty).Trim
            End If
        End If
        Return String.Empty
    End Function

    Public Function getYear() As String
        ''Year
        'tt0800320 <div id="tn15title">' \n '<h1>Clash of the Titans <span>(<a href="/year/2010/">2010</a>) <span class
        'tt0655454 <div id="tn15title">' \n '<h1>"Mystery Science Theater 3000" <span><em>The Beast of Yucca Flats</em> (1995)</span><span class="title-extra"></span></h1>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "<div id=""tn15title"">\s*<\w+>(?<title>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)<\w+>\s*(<em>(?<em>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)</em>\s*)?(\((<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>)?(?<year>\d{4}))?", RegexOptions.IgnoreCase)
        Return Regex.Match(iSource, "h1>((?<!</em>\s*\(|\(<a href|<h1).)+(=((?<!>).)+)?\s*(?<year>\d{4})", RegexOptions.IgnoreCase).Groups("year").Value
    End Function

    Public Function getAkas() As String
        ''AlsoKnownAs 
        'tt1055369 <div class="info"><h5>Also Known As:</h5><div class="info-content">       "Transformers: Revenge of the Fallen (IMAX DMR version)" - Hong Kong <em>(English title)</em> <em>(IMAX version)</em><br>"Transformers: Revenge of the Fallen - The IMAX Experience" - USA <em>(IMAX version)</em><br>"Transformer: Revenge" - Japan <em>(English title)</em><br>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Also Known As:</\w+>\s*<[\w\+\-\|/\\""'&=_ ]+>(\s*(?<akas>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\|][\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)(?<em>(\s*<em>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+\s*</em>(\s*<br ?/?>)*)+|))+", RegexOptions.IgnoreCase)
        'revision 1.31 : Match = Regex.Match(iSource, "Also Known As:(\s*<((?<!>).)+)+(?<akas>((?!(<a|<br|</div|<h)).)+<br ?/?>)+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Also Known As:(\s*<((?<!>).)+)+(?<akas>((?!(<a|<br|</div|<h)).)+<br ?/?>)+", RegexOptions.IgnoreCase)
        Dim strRegex As String = "(english title|international|original title|literal title|working)"
        Dim strAkas As String = String.Empty
        If getAltTitle() <> String.Empty Then
            strAkas = getMovTitle() & " - International (English title)" & vbNewLine
            strRegex = "(english title|international|working)"
        End If
        If Match.Success Then
            For gCtr As Integer = 0 To Match.Groups("akas").Captures.Count - 1
                If Regex.IsMatch(Match.Groups("akas").Captures.Item(gCtr).Value, strRegex, RegexOptions.IgnoreCase) Then
                    strAkas &= Regex.Replace(Match.Groups("akas").Captures.Item(gCtr).Value, "(<((?<!>).)+)+", String.Empty).Trim & vbNewLine
                End If
            Next
        End If
        Return charRef.cleanRef(strAkas.TrimEnd)
    End Function

    Public Function getCompany() As String
        ''Company
        '<div class="info"><h5>Company:</h5><div class="info-content"><a href="/company/co0040938/">DreamWorks SKG</a>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Company:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?\s*<[\w\+\-\.\?/\\""'#%&:;=_ ]+>(?<company>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>", RegexOptions.IgnoreCase)
        Return charRef.cleanRef(Regex.Match(iSource, "Company:(\s*<((?<!>).)+)+\s*(?<company>((?!</a>|</div|<h).)+)", RegexOptions.IgnoreCase).Groups("company").Value.TrimEnd)
    End Function

    Public Function getCertificate() As String
        ''Certificate
        'FILTER LIST REf. tt0800320, tt0120815
        '<div class="info"><h5>Certification:</h5><div class="info-content"><a href="us|pg_13">USA:PG-13</a> <i>(certificate #45430)</i> | <a href="ie|12a">Ireland:12A</a>  | <a href="kr|12">South Korea:12</a>  | <a href="jp|g">Japan:G</a>  | <a href="hk|iia">Hong Kong:IIA</a>  | <a href="fi|k_11">Finland:K-11</a> </div>
        'with quote <a href="/search/title?certificates=au|ma">Australia:MA</a> <i>(re-rating on appeal)</i>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Certification:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?(\s*<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>(?<cert>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>(?<i>(\s*<i>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+\s*</i>)|)(\s*\|)?)+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Certification:\s*(<((?<!>).)+)+(?<cert>((?!(</div>|\||<h)).)+(\|)?)+", RegexOptions.IgnoreCase)
        Dim strCert As String = String.Empty
        If Match.Success Then
            For gCtr As Integer = 0 To Match.Groups("cert").Captures.Count - 1
                strCert &= Regex.Replace(Match.Groups("cert").Captures.Item(gCtr).Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim & " | "
                If gCtr >= 4 Then : Exit For : End If
            Next
        End If
        Return charRef.cleanRef(strCert.TrimEnd(" "c, "|"c))
    End Function

    Public Function getCountry(Optional ByVal translate As Boolean = True) As String
        ''Country
        'tt1519640, countries '<div class="info"><h5>Country:</h5><div class="info-content"><a href="/Sections/Countries/USA/">USA</a> | <a href="/Sections/Countries/HongKong/">Hong Kong</a> | <a href="/Sections/Countries/China/">China</a></div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Country:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?(\s*<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>(?<country>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>(\s*\||)?)+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Country:(\s*<((?<!>).)+)+(?<country>((?!(</div>|\||<h)).)+(\|)?)+", RegexOptions.IgnoreCase)
        Dim strCountry As String = String.Empty
        If Match.Success Then
            For Each c As Capture In Match.Groups("country").Captures
                If translate = True Then : strCountry &= Regex.Replace(c.Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim & "/"
                Else : strCountry &= Regex.Replace(c.Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim.ToLower & "/" : End If
            Next
        End If
        Return strCountry.TrimEnd("/"c)
    End Function

    Public Function getGenre(Optional ByVal translate As Boolean = True) As String
        ''Genre
        '<div class="info"> <h5>Genre:</h5> <div class="info-content"> <a href="/Sections/Genres/Action/">Action</a> | <a href="/Sections/Genres/Adventure/">Adventure</a> | <a href="/Sections/Genres/Drama/">Drama</a> | <a href="/Sections/Genres/Fantasy/">Fantasy</a> <a class="tn15more inline" href="/title/tt0800320/keywords" onClick="(new Image()).src='/rg/title-tease/keywords/images/b.gif?link=/title/tt0800320/keywords';">See more</a>&nbsp;&raquo; </div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Genre:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?(\s*<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>(?<genre>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>(\s*\||)?)+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Genre:(\s*<((?<!>).)+)+(?<genre>((?!(<a class|\||</div|<h)).)+(\|)?)+", RegexOptions.IgnoreCase)
        Dim strGenre As String = String.Empty
        If Match.Success Then
            For Each c As Capture In Match.Groups("genre").Captures
                If translate = True Then : strGenre &= Regex.Replace(c.Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim & "/"
                Else : strGenre &= Regex.Replace(c.Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim.ToLower & "/" : End If
            Next
        End If
        Return strGenre.TrimEnd("/"c)
    End Function

    Public Function getDirector() As String
        ''Director
        'tt1519640, directors
        '<div id="director-info" class="info">vbNewLine<h5>Directors:</h5>vbNewLine<div class="info-content">vbNewLine<a href="/name/nm0151135/" onclick="(new Image()).src='/rg/directorlist/position-1/images/b.gif?link=name/nm0151135/';">Tony Chan</a><br/>vbNewLine
        '<a href="/name/nm0796102/" onclick="(new Image()).src='/rg/directorlist/position-2/images/b.gif?link=name/nm0796102/';">Wing Shya</a><br/>vbNewLine</div>vbNewLine</div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Directors?:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?(\s*<[\w\+\-\.\?\|/\(\)\\""'#%&:;=_ ]+>(?<director>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>(\s*<br ?/?>)?)+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Directors?:(\s*<((?<!>).)+)+(?<director>((?!<br ?/?|</div|<h).)+(<br ?/?>)?\s*)+", RegexOptions.IgnoreCase)
        Dim strDirector As String = String.Empty
        If Match.Success Then
            For Each c As Capture In Match.Groups("director").Captures
                strDirector &= Regex.Replace(c.Value, "(<((?<!>).)+)+", String.Empty).Trim & vbNewLine
            Next
        End If
        Return charRef.cleanRef(strDirector.TrimEnd)
    End Function

    Public Function getPremiereDate() As String
        ''Premiere Date
        '<div class="info">vbNewLine<h5>Release Date:</h5>vbNewLine<div class="info-content">vbNewLine24 June 2009 (Malaysia)vbNewLine<a class="tn15more inline" href="/title/tt1055369/releaseinfo" onClick="(new Image()).src='/rg/title-tease/releasedates/images/b.gif?link=/title/tt1055369/releaseinfo';">See more</a>&nbsp;&raquo;vbNewLine</div>vbNewLine</div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Release Date:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?\s*(?<date>\d{1,2})?\s*(?<month>\w+)\s*(?<year>\d{4})\s*(?<country>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*<", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Release Date:(\s*<((?<!>).)+)+\s*(?<date>\d{1,2})?\s*(?<month>\w+)\s*(?<year>\d{4})", RegexOptions.IgnoreCase)
        Dim strDate As String = String.Empty
        If Match.Success Then
            strDate = Match.Groups("year").Value & "年" & _
            Date.Parse("#01-" & Match.Groups("month").Value.ToString & "#").ToString("MM").TrimStart("0"c) & "月" & _
            IIf(Match.Groups("date").Success, Match.Groups("date").Value.TrimStart("0"c) & "日", "")
        End If
        Return strDate
    End Function

    Public Function getPremiereCountry() As String
        ''Premiere Country
        '<div class="info">vbNewLine<h5>Release Date:</h5>vbNewLine<div class="info-content">vbNewLine24 June 2009 (Malaysia)vbNewLine<a class="tn15more inline" href="/title/tt1055369/releaseinfo" onClick="(new Image()).src='/rg/title-tease/releasedates/images/b.gif?link=/title/tt1055369/releaseinfo';">See more</a>&nbsp;&raquo;vbNewLine</div>vbNewLine</div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Release Date:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?\s*(?<date>\d{1,2})?\s*(?<month>\w+)\s*(?<year>\d{4})\s*(?<country>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*<", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Release Date:(\s*<((?<!>).)+)+\s*((?<!\().)+(?<country>((?!\)|<a|</div|<h).)+)", RegexOptions.IgnoreCase)
        Dim strCountry As String = String.Empty
        If Match.Success Then
            strCountry = Match.Groups("country").Value.Trim.ToLower
        End If
        Return strCountry
    End Function

    Public Function getUserRating() As String
        ''UserRating
        'Not-IMDb Top 250 Rating
        'tt1519640, Match User Rating
        '<div class="starbar-meta">vbNewLine<b>6.0/10</b> vbNewLine  &nbsp;&nbsp;<a href="ratings" class="tn15more">91,322 votes</a>&nbsp;&raquo;vbNewLine</div>vbNewLine</div>
        'tt0010030, Empty User Rating
        '<div class="starbar-meta">vbNewLine<small>(awaiting 5 votes)</small>vbNewLine</div>
        'tt1375666, IMDb Top 250 Chart
        '<div class="starbar-special">vbNewLine<a href="/chart/top?tt1375666">Top 250: #6</a>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "<div class=""starbar\-meta"">\s*<\w+>\s*(?<rate>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>(\s*<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>(?<votes>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+))?\s*</\w+>", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "starbar\-meta"">\s*<\w+>(?<rating>((?!</\w+>).)+)</\w+>(\s*<a((?<!>).)+(?<votes>((?!</a|</div|<h).)+))?", RegexOptions.IgnoreCase)
        Dim strRate As String = String.Empty
        If Match.Success Then
            strRate = Match.Groups("rating").Value.Trim & _
            IIf(Match.Groups("votes").Success, " (" & Match.Groups("votes").Value.Trim & ")", "")
            'revision 1.02 : Match = Regex.Match(httpImdbSource, "<div class=""starbar\-special"">\s*<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>\s*(?<top250>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>", RegexOptions.IgnoreCase)
            Match = Regex.Match(iSource, "starbar\-special"">\s*<a((?<!>).)+(?<top250>((?!</a|</div|<h).)+)", RegexOptions.IgnoreCase)
            If Match.Groups("top250").Success Then
                strRate &= " " & Match.Groups("top250").Value.Trim
            End If
        End If
        Return strRate
    End Function

    Public Function getLanguage(Optional ByVal translate As Boolean = True) As String
        ''Language
        'tt0120815, Saving Private Ryan
        '<div class="info"><h5>Language:</h5><div class="info-content"><a href="/Sections/Languages/English/">English</a> | <a href="/Sections/Languages/French/">French</a> | <a href="/Sections/Languages/German/">German</a> | <a href="/Sections/Languages/Czech/">Czech</a></div></div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Language:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?(\s*<[\w\+\-\.\|\?/\\""'#%&:;=_ ]+>(?<language>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)\s*</\w+>(\s*\|)?)+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Language:(\s*<((?<!>).)+)+(?<lua>((?!(</div>|\||<h)).)+(\|)?)+", RegexOptions.IgnoreCase)
        Dim strLanguage As String = String.Empty
        If Match.Success Then
            For Each c As Capture In Match.Groups("lua").Captures
                If translate = True Then : strLanguage &= Regex.Replace(c.Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim & "/"
                Else : strLanguage &= Regex.Replace(c.Value, "((<((?<!>).)+)+|\|)", String.Empty).Trim.ToLower & "/" : End If
            Next
        End If
        Return strLanguage.Trim("/"c)
    End Function

    Public Function getCast() As String
        ''Cast
        'tt0800320, default case
        'tt0800320, actor as multiple characters
        'tt0010030, spliter in the middle
        'tt0010030, character name not available
        'tt1519640, character field without hyperlink
        'tt0120815, character field without hyperlink

        '1L
        '<td class="nm"><a href="/name/nm0516001/" onclick="(new Image()).src=' .. /name/nm0516001/';">Harold Lloyd</a></td>
        '1L ... 1
        '<td class="nm"><a href="/name/nm0874866/" onclick="(new Image()).src=' .. /name/nm0874866/';">Nicholas Tse</a></td><td class="ddd"> ... </td><td class="char">Ah Wai</td></tr>
        '1L     1
        '<td class="nm"><a href="/name/nm0430865/" onclick="(new Image()).src=' .. /name/nm0430865/';">Margaret Joslin</a></td><td class="ddd"></td><td class="char"> (as Margaret Joslyn)</td></tr>
        '1L ... 1L
        '<td class="nm"><a href="/name/nm0941777/" onclick="(new Image()).src=' .. /name/nm0941777/';">Sam Worthington</a></td><td class="ddd"> ... </td><td class="char"><a href="/character/ch0043859/">Perseus</a></td></tr>  '
        '1L ... 1L (as ooxx)
        '<td class="char"><a href="/character/ch0002093/">Cpl. Henderson</a> (as Maximilian Martini)</td></tr>
        '1L ... 2L
        '<td class="nm"><a href="/name/nm0002076/" onclick="(new Image()).src=' .. /name/nm0002076/';">Jason Flemyng</a></td><td class="ddd"> ... </td><td class="char"><a href="/character/ch0027406/">Calibos</a> / <a href="/character/ch0144056/">Acrisius</a></td></tr>

        '.tester()
        'header <td class=""nm"">
        'header <a href="/name/nm0941777/" onclick="(new Image()).src='...';">
        'header Harold Lloyd</a></td>
        'op-middle <td class="ddd">
        'op(-op - middle) ' ... '
        'op-middle </td>
        'op-rp-footer <td class="char">
        'op-rp-op-footer <a href="">
        'op-rp-footer Perseus
        'op-rp-op-footer </a>
        'op-rp-op-footer (as ooxx)
        'op-rp-footer </td></tr>
        'op-rp-op-footer " / "

        'revision 1.02 : > ========= BACKUP ===========
        'Dim Matches As MatchCollection
        'Dim strCast As New StringBuilder
        'Matches = Regex.Matches(httpImdbSource, "<td class=""nm"">\s*<[\w\+\-\.\?\|/\(\)\\""'#%&:;=_ ]+>\s*(?<nm>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)(\s*</\w+>)+(\s*<[\w\+\-\|/\\""'&=_ ]+>\s*[""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]*\s*</\w+>(\s*<[\w\+\-\|/\\""'&=_ ]+>(\s*<[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>)?\s*(?<char>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+)(\s*</\w+>)+\s*(?<as>[\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]+|)(\s*/)?)+)?", RegexOptions.IgnoreCase)
        'If Matches.Count >= 1 Then
        '    For groupCtr As Integer = 0 To Matches.Count - 1
        '        strCast.Append(Matches.Item(groupCtr).Groups("nm").Value.TrimEnd)
        '        For captureCtr As Integer = 0 To Matches.Item(groupCtr).Groups("char").Captures.Count - 1
        '            If captureCtr = 0 Then
        '                strCast.Append(" ... " & Matches.Item(groupCtr).Groups("char").Captures.Item(captureCtr).Value.TrimEnd & Matches.Item(groupCtr).Groups("as").Captures.Item(captureCtr).Value.TrimEnd)
        '            Else
        '                strCast.Append(" / " & Matches.Item(groupCtr).Groups("char").Captures.Item(captureCtr).Value.TrimEnd & Matches.Item(groupCtr).Groups("as").Captures.Item(captureCtr).Value.TrimEnd)
        '            End If
        '        Next
        '        strCast.AppendLine()
        '        If groupCtr >= 14 Then
        '            Exit For
        '        End If
        '    Next
        '    info.Cast = charRef.addRef(strCast.ToString.TrimEnd)
        'End If
        'revision 1.02 : < ========= BACKUP ===========
        Dim Match As Match = Regex.Match(iSource, "table class=""cast"">\s*(?<cast>((?!</tr|</table|</div).)+</tr>\s*)+", RegexOptions.IgnoreCase)
        Dim strCast As String = String.Empty
        If Match.Success Then
            Dim gCtr As Integer = 0
            For Each c As Capture In Match.Groups("cast").Captures
                If Not Regex.Match(Match.Groups("cast").Captures.Item(gCtr).Value, "^<tr>").Success Then
                    strCast &= Regex.Replace(Match.Groups("cast").Captures.Item(gCtr).Value, "(<((?<!>).)+)+", String.Empty).Trim & vbNewLine
                End If
                If gCtr >= 14 Then : Exit For : End If
                gCtr += 1
            Next
        End If
        Return charRef.cleanRef(strCast.TrimEnd)
    End Function

    Public Function getPlot() As String
        ''Plot
        'tt1385867, Short Summary
        '<div class="info">vbNewLine<h5>Plot:</h5>vbNewLine<div class="info-content">vbNewLine A comedy about ... gangster. | <a class="tn15more inline" href="synopsis">Add synopsis</a>&nbsp;&raquo;vbNewLine</div>vbNewLine</div>
        'tt0120815, Full Summary
        '<div class="info">vbNewLine<h5>Plot:</h5>vbNewLine<div class="info-content">vbNewLine Following  ... in action. <a class="tn15more inline" href="/title/tt0120815/plotsummary" onClick="(new Image()).src='/rg/title-tease/plotsummary/images/b.gif?link=/title/tt0120815/plotsummary';">Full summary</a>&nbsp;&raquo; | <a class="tn15more inline" href="synopsis">Full synopsis</a>&nbsp;&raquo;vbNewLine</div>vbNewLine</div>
        'tt0800320, Full Synopsis
        '<div class="info">vbNewLine<h5>Plot:</h5>vbNewLineThe mortal son of the god Zeus embarks on a perilous journey to stop the underworld and its minions from spreading their evil to Earth as well as the heavens. | <a class="tn15more inline" href="synopsis">Full synopsis</a>&nbsp;&raquo; vbNewLine </div>
        'tt1094249, HyperLink Plot
        '<div class="info">vbNewLine<h5>Plot:</h5>vbNewLine<div class="info-content">vbNewLine A short prologue of one heartbreaking history of love and the prologue of the travel told in <a href="/title/tt0838221/">The Darjeeling Limited</a> (2007). <a class="tn15more inline" href="/title/tt1094249/plotsummary" onClick="(new Image()).src='/rg/title-tease/plotsummary/images/b.gif?link=/title/tt1094249/plotsummary';">Full summary</a>&nbsp;&raquo; | <a class="tn15more inline" href="synopsis">Add synopsis</a>&nbsp;&raquo; vbNewLine </div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Plot:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?\s*(?<plot>([\w""'`!#$%&*,:;=@~/\\_\+\-\.\?\^\(\)\[\]\{\}\| ]|<a href[\w\+\-\.\?\|/\\""'#%&:;=_ ]+>|</a>)+)<", RegexOptions.IgnoreCase)
        'revision 1.02 : info.Plot = charRef.addRef(Regex.Replace(Match.Groups("plot").Value, "<[\w\+\-\.\?\|/\\""'#%&:;=_ ]*>", "").Trim(" "c, "|"c))
        Return charRef.cleanRef(Regex.Replace( _
                           Regex.Match(iSource, "Plot:(\s*<((?<!>).)+)+\s*(?<plot>((?!\|? <a class|</div|<h).)+)", _
                                       RegexOptions.IgnoreCase).Groups("plot").Value, "<((?<!>).)+", String.Empty).Trim)
    End Function

    Public Function getRuntime() As String
        ''Runtime
        '<div class="info"><h5>Runtime:</h5><div class="info-content">150 min </div></div>
        'tt1519640, Non-Digit Runtime
        '<div class="info"><h5>Runtime:</h5><div class="info-content">Hong Kong:93 min  | USA:93 min </div></div>
        'revision 1.02 : Match = Regex.Match(httpImdbSource, "Runtime:</\w+>(\s*<div[\w\+\-\|/\\""'&=_ ]+>)?\s*(?<runtime>\d{1,3})\s*\w+", RegexOptions.IgnoreCase)
        Dim Match As Match = Regex.Match(iSource, "Runtime:(\s*<((?<!>).)+)+(?<dur>\d+|((?!</div|<h).)+)", RegexOptions.IgnoreCase)
        Dim strDur As String = String.Empty
        If Match.Success Then
            If Regex.IsMatch(Match.Groups("dur").Value, "^\d+$") Then
                strDur = Match.Groups("dur").Value
            ElseIf Regex.IsMatch(Match.Groups("dur").Value, "\d+") Then
                strDur = Regex.Match(Match.Groups("dur").Value, "\d+").Value
            End If
        End If
        Return strDur
    End Function
#End Region

#Region "> ======== CHARACTER REFERENCE ========"
    Protected Class charRef
        Shared Function cleanRef(ByVal Source As String) As String
            '' convert numeric character reference & character entity reference to character
            'numeric code references in Decimal (base-10)
            cleanRef = Regex.Replace(Source, "&#(?<ncr>\d+);", AddressOf ncrDec)
            'numeric code references in Hexadecimal (base-16)
            cleanRef = Regex.Replace(cleanRef, "&#x(?<ncr>[a-fA-F0-9]+);", AddressOf ncrHex)
            'character entity references
            cleanRef = Regex.Replace(cleanRef, "&(?<cer>\w+);", AddressOf ceRef)
            'remove html escape characters
            cleanRef = Regex.Replace(cleanRef, "\\[^\\]", AddressOf espChr)
            Return cleanRef
        End Function

        Private Shared Function ncrDec(ByVal m As Match) As Char
            Return Convert.ToChar(Integer.Parse(m.Groups("ncr").Value))
        End Function

        Private Shared Function ncrHex(ByVal m As Match) As Char
            Return Convert.ToChar(Convert.ToInt32(m.Groups("ncr").Value, 16))
        End Function

        Private Shared Function espChr(ByVal m As Match) As Char
            Return m.Value.Remove(0, 1)
        End Function

        Private Shared Function ceRef(ByVal m As Match) As Char
            Select Case m.Groups("cer").Value
                'common character entity references
                Case "quot" : Return """"
                Case "amp" : Return "&"
                Case "apos" : Return "'"
                Case "lt" : Return "<"
                Case "gt" : Return ">"
                Case "nbsp" : Return " "
                'rare character entity references
                Case "iexcl" : Return "¡"
                Case "cent" : Return "¢"
                Case "pound" : Return "£"
                Case "curren" : Return "¤"
                Case "yen" : Return "¥"
                Case "brvbar" : Return "¦"
                Case "sect" : Return "§"
                Case "uml" : Return "¨"
                Case "copy" : Return "©"
                Case "ordf" : Return "ª"
                Case "laquo" : Return "«"
                Case "not" : Return "¬"
                Case "shy" : Return ""
                Case "reg" : Return "®"
                Case "macr" : Return "¯"
                Case "deg" : Return "°"
                Case "plusmn" : Return "±"
                Case "sup2" : Return "²"
                Case "sup3" : Return "³"
                Case "acute" : Return "´"
                Case "micro" : Return "µ"
                Case "para" : Return "¶"
                Case "middot" : Return "·"
                Case "cedil" : Return "¸"
                Case "sup1" : Return "¹"
                Case "ordm" : Return "º"
                Case "raquo" : Return "»"
                Case "frac14" : Return "¼"
                Case "frac12" : Return "½"
                Case "frac34" : Return "¾"
                Case "iquest" : Return "¿"
                Case "Agrave" : Return "À"
                Case "Aacute" : Return "Á"
                Case "Acirc" : Return "Â"
                Case "Atilde" : Return "Ã"
                Case "Auml" : Return "Ä"
                Case "Aring" : Return "Å"
                Case "AElig" : Return "Æ"
                Case "Ccedil" : Return "Ç"
                Case "Egrave" : Return "È"
                Case "Eacute" : Return "É"
                Case "Ecirc" : Return "Ê"
                Case "Euml" : Return "Ë"
                Case "Igrave" : Return "Ì"
                Case "Iacute" : Return "Í"
                Case "Icirc" : Return "Î"
                Case "Iuml" : Return "Ï"
                Case "ETH" : Return "Ð"
                Case "Ntilde" : Return "Ñ"
                Case "Ograve" : Return "Ò"
                Case "Oacute" : Return "Ó"
                Case "Ocirc" : Return "Ô"
                Case "Otilde" : Return "Õ"
                Case "Ouml" : Return "Ö"
                Case "times" : Return "×"
                Case "Oslash" : Return "Ø"
                Case "Ugrave" : Return "Ù"
                Case "Uacute" : Return "Ú"
                Case "Ucirc" : Return "Û"
                Case "Uuml" : Return "Ü"
                Case "Yacute" : Return "Ý"
                Case "THORN" : Return "Þ"
                Case "szlig" : Return "ß"
                Case "agrave" : Return "à"
                Case "aacute" : Return "á"
                Case "acirc" : Return "â"
                Case "atilde" : Return "ã"
                Case "auml" : Return "ä"
                Case "aring" : Return "å"
                Case "aelig" : Return "æ"
                Case "ccedil" : Return "ç"
                Case "egrave" : Return "è"
                Case "eacute" : Return "é"
                Case "ecirc" : Return "ê"
                Case "euml" : Return "ë"
                Case "igrave" : Return "ì"
                Case "iacute" : Return "í"
                Case "icirc" : Return "î"
                Case "iuml" : Return "ï"
                Case "eth" : Return "ð"
                Case "ntilde" : Return "ñ"
                Case "ograve" : Return "ò"
                Case "oacute" : Return "ó"
                Case "ocirc" : Return "ô"
                Case "otilde" : Return "õ"
                Case "ouml" : Return "ö"
                Case "divide" : Return "÷"
                Case "oslash" : Return "ø"
                Case "ugrave" : Return "ù"
                Case "uacute" : Return "ú"
                Case "ucirc" : Return "û"
                Case "uuml" : Return "ü"
                Case "yacute" : Return "ý"
                Case "thorn" : Return "þ"
                Case "yuml" : Return "ÿ"
                Case "oelig" : Return "œ"
                Case "oelig" : Return "œ"
                Case "scaron" : Return "š"
                Case "scaron" : Return "š"
                Case "yuml" : Return "ÿ"
                Case "fnof" : Return "ƒ"
                Case "circ" : Return "ˆ"
                Case "tilde" : Return "˜"
                Case "Alpha" : Return "Α"
                Case "Beta" : Return "Β"
                Case "Gamma" : Return "Γ"
                Case "Delta" : Return "Δ"
                Case "Epsilon" : Return "Ε"
                Case "Zeta" : Return "Ζ"
                Case "Eta" : Return "Η"
                Case "Theta" : Return "Θ"
                Case "Iota" : Return "Ι"
                Case "Kappa" : Return "Κ"
                Case "Lambda" : Return "Λ"
                Case "Mu" : Return "Μ"
                Case "Nu" : Return "Ν"
                Case "Xi" : Return "Ξ"
                Case "Omicron" : Return "Ο"
                Case "Pi" : Return "Π"
                Case "Rho" : Return "Ρ"
                Case "Sigma" : Return "Σ"
                Case "Tau" : Return "Τ"
                Case "Upsilon" : Return "Υ"
                Case "Phi" : Return "Φ"
                Case "Chi" : Return "Χ"
                Case "Psi" : Return "Ψ"
                Case "Omega" : Return "Ω"
                Case "alpha" : Return "α"
                Case "beta" : Return "β"
                Case "gamma" : Return "γ"
                Case "delta" : Return "δ"
                Case "epsilon" : Return "ε"
                Case "zeta" : Return "ζ"
                Case "eta" : Return "η"
                Case "theta" : Return "θ"
                Case "iota" : Return "ι"
                Case "kappa" : Return "κ"
                Case "lambda" : Return "λ"
                Case "mu" : Return "μ"
                Case "nu" : Return "ν"
                Case "xi" : Return "ξ"
                Case "omicron" : Return "ο"
                Case "pi" : Return "π"
                Case "rho" : Return "ρ"
                Case "sigmaf" : Return "ς"
                Case "sigma" : Return "σ"
                Case "tau" : Return "τ"
                Case "upsilon" : Return "υ"
                Case "phi" : Return "φ"
                Case "chi" : Return "χ"
                Case "psi" : Return "ψ"
                Case "omega" : Return "ω"
                Case "thetasym" : Return "ϑ"
                Case "upsih" : Return "ϒ"
                Case "piv" : Return "ϖ"
                Case "ensp" : Return " "
                Case "emsp" : Return " "
                Case "zwnj" : Return ""
                Case "zwj" : Return ""
                Case "lrm" : Return ""
                Case "rlm" : Return ""
                Case "ndash" : Return "–"
                Case "mdash" : Return "—"
                Case "lsquo" : Return "‘"
                Case "rsquo" : Return "’"
                Case "sbquo" : Return "‚"
                Case "ldquo" : Return """"
                Case "rdquo" : Return """"
                Case "bdquo" : Return "„"
                Case "dagger" : Return "†"
                Case "Dagger" : Return "‡"
                Case "bull" : Return "•"
                Case "hellip" : Return "…"
                Case "permil" : Return "‰"
                Case "prime" : Return "′"
                Case "Prime" : Return "″"
                Case "lsaquo" : Return "‹"
                Case "rsaquo" : Return "›"
                Case "oline" : Return "‾"
                Case "frasl" : Return "⁄"
                Case "euro" : Return "€"
                Case "image" : Return "ℑ"
                Case "weierp" : Return "℘"
                Case "real" : Return "ℜ"
                Case "trade" : Return "™"
                Case "alefsym" : Return "ℵ"
                Case "larr" : Return "←"
                Case "uarr" : Return "↑"
                Case "rarr" : Return "→"
                Case "darr" : Return "↓"
                Case "harr" : Return "↔"
                Case "crarr" : Return "↵"
                Case "lArr" : Return "⇐"
                Case "uArr" : Return "⇑"
                Case "rArr" : Return "⇒"
                Case "dArr" : Return "⇓"
                Case "hArr" : Return "⇔"
                Case "forall" : Return "∀"
                Case "part" : Return "∂"
                Case "exist" : Return "∃"
                Case "empty" : Return "∅"
                Case "nabla" : Return "∇"
                Case "isin" : Return "∈"
                Case "notin" : Return "∉"
                Case "ni" : Return "∋"
                Case "prod" : Return "∏"
                Case "sum" : Return "∑"
                Case "minus" : Return "−"
                Case "lowast" : Return "∗"
                Case "radic" : Return "√"
                Case "prop" : Return "∝"
                Case "infin" : Return "∞"
                Case "ang" : Return "∠"
                Case "and" : Return "∧"
                Case "or" : Return "∨"
                Case "cap" : Return "∩"
                Case "cup" : Return "∪"
                Case "int" : Return "∫"
                Case "there4" : Return "∴"
                Case "sim" : Return "∼"
                Case "cong" : Return "≅"
                Case "asymp" : Return "≈"
                Case "ne" : Return "≠"
                Case "equiv" : Return "≡"
                Case "le" : Return "≤"
                Case "ge" : Return "≥"
                Case "sub" : Return "⊂"
                Case "sup" : Return "⊃"
                Case "nsub" : Return "⊄"
                Case "sube" : Return "⊆"
                Case "supe" : Return "⊇"
                Case "oplus" : Return "⊕"
                Case "otimes" : Return "⊗"
                Case "perp" : Return "⊥"
                Case "sdot" : Return "⋅"
                Case "lceil" : Return "⌈"
                Case "rceil" : Return "⌉"
                Case "lfloor" : Return "⌊"
                Case "rfloor" : Return "⌋"
                Case "lang" : Return "〈"
                Case "rang" : Return "〉"
                Case "loz" : Return "◊"
                Case "spades" : Return "♠"
                Case "clubs" : Return "♣"
                Case "hearts" : Return "♥"
                Case "diams" : Return "♦"
                Case Else : Return ""
            End Select
        End Function
    End Class
#End Region

End Class
