Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class Sucursal
    Implements IDisposable
    Public Const FECHA_NULA As Date = #12/31/1599#

    Private cn As OracleConnection
    Private da1 As OracleDataAdapter 'Sucursal
    Private da2 As OracleDataAdapter 'Sucursal Entregas
    Private dt1 As New DataTable 'Sucursal
    Private dt2 As New DataTable 'Sucursal Entregas
    Private dr1 As DataRow  'Sucursal
    Private dr2 As DataRow  'Sucursal Entregas

    Private _Pais As Pais = Nothing
    Private disposedValue As Boolean = False ' Para detectar llamadas redundantes

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()

        da1.FillSchema(dt1, SchemaType.Mapped)
        da2.FillSchema(dt2, SchemaType.Mapped)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Tercero As String, ByVal Sucursal As String)
        Me.New(cn)
        Abrir(Tercero, Sucursal)
    End Sub

    'SUB
    Public Function Abrir(ByVal Tercero As String, ByVal Sucursal As String) As Boolean
        _Pais = Nothing

        dt1.Clear()
        dt2.Clear()
        da1.SelectCommand.Parameters("bpanum_0").Value = Tercero
        da1.SelectCommand.Parameters("bpaadd_0").Value = Sucursal
        da1.Fill(dt1)
        If dt1.Rows.Count > 0 Then
            dr1 = dt1.Rows(0)
        End If

        da2.SelectCommand.Parameters("bpcnum_0").Value = Tercero
        da2.SelectCommand.Parameters("bpaadd_0").Value = Sucursal
        da2.Fill(dt2)
        If dt2.Rows.Count = 1 Then dr2 = dt2.Rows(0)
        Return dt1.Rows.Count > 0

    End Function
    Friend Function Abrir(ByVal drSuc As DataRow, ByVal drEnt As DataRow) As Boolean
        _Pais = Nothing

        'Sucursal
        dt1 = Nothing
        dt1 = drSuc.Table.Clone

        dr1 = dt1.NewRow

        For i = 0 To dt1.Columns.Count - 1
            dr1(i) = drSuc(i)
        Next

        dt1.Rows.Add(dr1)
        dt1.AcceptChanges()

        'Sucursal Entrega
        If drEnt IsNot Nothing Then
            dt2 = Nothing
            dt2 = drEnt.Table.Clone

            dr2 = dt2.NewRow

            For i = 0 To dt2.Columns.Count - 1
                dr2(i) = drEnt(i)
            Next

            dt2.Rows.Add(dr2)
            dt2.AcceptChanges()

        End If

        Return True
    End Function
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM bpaddress WHERE bpanum_0 = :bpanum_0 AND bpaadd_0 = :bpaadd_0"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar)
        da1.SelectCommand.Parameters.Add("bpaadd_0", OracleType.VarChar)

        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        'da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand

        Sql = "UPDATE BPADDRESS SET BPATYP_0 = :BPATYP_0,BPANUM_0 = :BPANUM_0,BPAADD_0 = :BPAADD_0,BPADES_0 = :BPADES_0,BPAADDFLG_0 = :BPAADDFLG_0, "
        Sql &= "BPAADDLIG_0 = :BPAADDLIG_0,BPAADDLIG_1 = :BPAADDLIG_1,BPAADDLIG_2 = :BPAADDLIG_2,BPAADDNRO_0 = :BPAADDNRO_0,BPAGCBA_0 = :BPAGCBA_0,POSCOD_0 =:POSCOD_0, "
        Sql &= "CTY_0 = :CTY_0,SAT_0 = :SAT_0,CRY_0 = :CRY_0,CRYNAM_0 =:CRYNAM_0,TEL_0 =:TEL_0,FAX_0 =:FAX_0,WEB_0 =:WEB_0,EXTNUM_0 = :EXTNUM_0,EXPNUM_0 =:EXPNUM_0, "
        Sql &= "CREUSR_0 =:CREUSR_0,CREDAT_0 =:CREDAT_0,CRETIM_0 =:CRETIM_0,UPDUSR_0 =:UPDUSR_0,UPDDAT_0 =:UPDDAT_0,UPDTIM_0 = :UPDTIM_0,XENTFAC_0=:XENTFAC_0, "
        Sql &= "XMAILFC_0 =:XMAILFC_0,XBPASTA_0 =:XBPASTA_0,IRAM_OK_0 =:IRAM_OK_0,XPORTERO_0 =:XPORTERO_0,XPOR_ADD_0 =:XPOR_ADD_0,XPOR_TEL_0=:XPOR_TEL_0, "
        Sql &= "XPOR_MAIL_0 =:XPOR_MAIL_0,XTIPOSIST_0 =:XTIPOSIST_0,XHIDRANTE_0 =:XHIDRANTE_0,XTOMAS_0 =:XTOMAS_0,XESCLUSA_0 =:XESCLUSA_0,XBOMBAS_0 =:XBOMBAS_0, "
        Sql &= "XHIDRO_0 =:XHIDRO_0,XROCIADO_0 =:XROCIADO_0,XSISTOTRO_0 =:XSISTOTRO_0,XINISERVI_0 =:XINISERVI_0, XPH_0 =:XPH_0, XPLANO_0 = :XPLANO_0, "
        Sql &= "XFILTRO_0 = :XFILTRO_0, XCOMBUST_0 = :XCOMBUST_0, XBOMBISTA_0 = :XBOMBISTA_0, XCURVA_0 = :XCURVA_0, XFINEG_0 = :XFINEG_0 "
        Sql &= "where bpanum_0 = :bpanum_0 and bpaadd_0 =  :bpaadd_0"
        da1.UpdateCommand = New OracleCommand(Sql, cn)

        With da1
            Parametro(.UpdateCommand, "BPATYP_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPANUM_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAADD_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPADES_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAADDFLG_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAADDLIG_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAADDLIG_1", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAADDLIG_2", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAADDNRO_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "BPAGCBA_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "POSCOD_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "CTY_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "SAT_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "CRY_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "CRYNAM_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "TEL_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "FAX_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "WEB_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "EXTNUM_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "EXPNUM_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "CREUSR_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "CREDAT_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "CRETIM_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "UPDUSR_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "UPDDAT_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "UPDTIM_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XENTFAC_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XMAILFC_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XBPASTA_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "IRAM_OK_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XPORTERO_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XPOR_ADD_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XPOR_TEL_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XPOR_MAIL_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XTIPOSIST_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XHIDRANTE_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XTOMAS_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XESCLUSA_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XBOMBAS_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XHIDRO_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XROCIADO_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XSISTOTRO_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XINISERVI_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XPH_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XPLANO_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XFILTRO_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XCOMBUST_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XBOMBISTA_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XCURVA_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "XFINEG_0", OracleType.VarChar, DataRowVersion.Current)
        End With

        Sql = "SELECT * FROM bpdlvcust WHERE bpcnum_0 = :bpcnum_0 AND bpaadd_0 = :bpaadd_0"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar)
        da2.SelectCommand.Parameters.Add("bpaadd_0", OracleType.VarChar)
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand

    End Sub
    Public Sub Grabar()
        Try
            da1.Update(dt1)
            da2.Update(dt2)
        Catch ex As Exception

        End Try
    End Sub
    Public Sub Nuevo(ByVal Cliente As String)
        _Pais = Nothing

        Dim dr As DataRow
        dr = dt1.NewRow
        dr("BPATYP_0") = 1
        dr("BPANUM_0") = Cliente
        dr("BPAADD_0") = " "
        dr("BPADES_0") = " "
        dr("BPAADDFLG_0") = 0
        dr("BPAADDLIG_0") = " "
        dr("BPAADDLIG_1") = " "
        dr("BPAADDLIG_2") = " "
        dr("BPAADDNRO_0") = 0
        dr("BPAGCBA_0") = " "
        dr("POSCOD_0") = " "
        dr("CTY_0") = " "
        dr("SAT_0") = " "
        dr("CRY_0") = "AR"
        dr("CRYNAM_0") = "Argentina"
        dr("TEL_0") = " "
        dr("FAX_0") = " "
        dr("WEB_0") = " "
        dr("EXTNUM_0") = " "
        dr("EXPNUM_0") = 0
        dr("CREUSR_0") = " "
        dr("CREDAT_0") = Date.Today
        dr("CRETIM_0") = Today.Hour * 3600 + Today.Minute * 60
        dr("UPDUSR_0") = " "
        dr("UPDDAT_0") = FECHA_NULA
        dr("UPDTIM_0") = 0
        dr("XENTFAC_0") = 0
        dr("XMAILFC_0") = " "
        dr("XBPASTA_0") = 0
        dr("IRAM_OK_0") = 0
        dr("XPORTERO_0") = " "
        dr("XPOR_ADD_0") = " "
        dr("XPOR_TEL_0") = " "
        dr("XPOR_MAIL_0") = " "
        'Desarrollo de guido
        dr("xtiposist_0") = 0
        dr("xhidrante_0") = 0
        dr("xtomas_0") = 0
        dr("xesclusa_0") = 0
        dr("xph_0") = 0
        dr("xbombas_0") = 0
        dr("xhidro_0") = 0
        dr("xrociado_0") = 0
        dr("xsistotro_0") = 0
        dr("xiniservi_0") = #12/31/1599#
        dr("XPLANO_0") = 0
        dr("XFILTRO_0") = 0
        dr("XCOMBUST_0") = 0
        dr("XBOMBISTA_0") = 0
        dr("XCURVA_0") = 0
        dr("XFINEG_0") = " "

        dt1.Rows.Add(dr)

        dr1 = dt1.Rows(0)
        'dr2 = dt2.Rows(0)

    End Sub
    Private Sub NuevaSucursalEntrega() 'ByVal Cliente As String)
        Dim dr As DataRow = dt2.NewRow
        dr = dt2.NewRow
        dr("BPCNUM_0") = Me.Codigo
        dr("BPAADD_0") = " "
        dr("BPDNAM_0") = " "
        dr("BPDNAM_1") = " "
        dr("STOFCY_0") = "D01"
        dr("RCPFCY_0") = " "
        dr("BPCLOC_0") = "CLI"
        dr("SCOLOC_0") = " "
        dr("LAN_0") = "SPA"
        dr("BPTNUM_0") = " "
        dr("MDL_0") = 1
        dr("EECICT_0") = " "
        dr("EECLOC_0") = 0
        dr("DRN_0") = 0
        dr("DLVPIO_0") = 1
        dr("DAYLTI_0") = 0
        dr("UVYDAY1_0") = 2
        dr("UVYDAY2_0") = 2
        dr("UVYDAY3_0") = 2
        dr("UVYDAY4_0") = 2
        dr("UVYDAY5_0") = 2
        dr("UVYDAY6_0") = 1
        dr("UVYDAY7_0") = 1
        dr("UVYCOD_0") = " "
        dr("EECINCRAT_0") = 0
        dr("REP_0") = " "
        dr("REP_1") = " "
        dr("REP_2") = " "
        dr("GEOCOD_0") = " "
        dr("INSCTYFLG_0") = " "
        dr("TAXEXN_0") = " "
        dr("BPDEXNFLG_0") = " "
        dr("PRPTEX_0") = " "
        dr("DLVTEX_0") = " "
        dr("NPRFLG_0") = 2
        dr("NDEFLG_0") = 2
        dr("EXPNUM_0") = 0
        dr("CREUSR_0") = " "
        dr("CREDAT_0") = FECHA_NULA
        dr("UPDUSR_0") = " "
        dr("UPDDAT_0") = Today.Date
        dr("YHDESDE1_0") = " "
        dr("YHDESDE2_0") = " "
        dr("YHHASTA1_0") = " "
        dr("YHHASTA2_0") = " 2"
        dr("ENAFLG_0") = 2
        dr("FLGBAJA_0") = 0
        dr("TMPESP_0") = 0
        dr("tranum_0") = " "
        dr("xmunro_0") = 0
        dt2.Rows.Add(dr)

        dr2 = dt2.Rows(0)

    End Sub
    Public Function Parque() As ParqueCollection
        Dim p As New ParqueCollection(cn)
        p.AbrirParqueCliente(Codigo, Sucursal)
        Return p
    End Function
    'FUNCTION

    'PROPERTY
    Public ReadOnly Property Tercero() As String
        Get
            Return dr1("bpanum_0").ToString
        End Get
    End Property
    Public Property Sucursal() As String
        Get
            If dt1.Rows.Count > 0 Then
                Return dr1("bpaadd_0").ToString
            Else
                Return ""
            End If

        End Get
        Set(ByVal value As String)
            If dt1.Rows.Count > 0 Then
                dr1.BeginEdit()
                dr1("bpaadd_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr1.EndEdit()
            End If
            If dt2.Rows.Count > 0 Then
                dr2.BeginEdit()
                dr2("bpaadd_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr2.EndEdit()
            End If
        End Set
    End Property
    Public Property Codigo() As String
        Get
            If dt1.Rows.Count > 0 Then
                Return dr1("bpanum_0").ToString
            Else
                Return ""
            End If

        End Get
        Set(ByVal value As String)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                dr.BeginEdit()
                dr("bpanum_0") = value
                dr.EndEdit()
            End If
            If dt2.Rows.Count > 0 Then
                Dim dr As DataRow = dt2.Rows(0)
                dr.BeginEdit()
                dr("bpcnum_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property CodigoNombre() As String
        Get
            Dim t As String

            t = Me.Sucursal _
              & " - " _
              & Me.Direccion

            Return t
        End Get
    End Property
    Public Property Nombre() As String
        Get
            Return dr1("bpades_0").ToString
        End Get
        Set(ByVal value As String)
            dr1.BeginEdit()
            dr1("bpades_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property NombreCliente() As String
        Get
            Dim txt As String = ""

            If dt2.Rows.Count > 0 Then
                txt = dr2("bpdnam_0").ToString.Trim
            End If

            Return txt

        End Get
        Set(ByVal value As String)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            Dim txt As String = IIf(value.Trim = "", " ", value.ToUpper).ToString
            dr2("bpdnam_0") = txt
            dr2.EndEdit()
        End Set
    End Property
    Public Property Direccion() As String
        Get
            Return dr1("bpaaddlig_0").ToString
        End Get
        Set(ByVal value As String)
            dr1.BeginEdit()
            dr1("bpaaddlig_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Direccion2() As String
        Get
            Return dr1("bpaaddlig_1").ToString
        End Get
    End Property
    Public ReadOnly Property Direccion3() As String
        Get
            Return dr1("bpaaddlig_2").ToString
        End Get
    End Property
    Public Property CodigoPostal() As String
        Get
            Return dr1("poscod_0").ToString
        End Get
        Set(ByVal value As String)
            dr1.BeginEdit()
            dr1("poscod_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Ciudad() As String
        Get
            Return dr1("cty_0").ToString
        End Get
        Set(ByVal value As String)
            dr1.BeginEdit()
            dr1("cty_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Provincia() As String
        Get
            Return dr1("sat_0").ToString
        End Get
        Set(ByVal value As String)
            dr1.EndEdit()
            dr1("sat_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public ReadOnly Property ProvinciaNombre() As String
        Get
            Select Case Me.Provincia
                Case "CFE"
                    Return "CAPITAL FEDERAL"
                Case "BUE"
                    Return "BUENOS AIRES"
                Case "CBA"
                    Return "CORDOBA"
                Case "CHA"
                    Return "CHACO"
                Case "CHU"
                    Return "CHUBUT"
                Case "COR"
                    Return "CORRIENTES"
                Case "CTC"
                    Return "CATAMARCA"
                Case "ERI"
                    Return "ENTRE RIOS"
                Case "FMA"
                    Return "FORMOSA"
                Case "JJY"
                    Return "JUJUY"
                Case "LPA"
                    Return "LA PAMPA"
                Case "LRJ"
                    Return "LA RIOJA"
                Case "MDZ"
                    Return "MENDOZA"
                Case "MIS"
                    Return "MISIONES"
                Case "NQN"
                    Return "NEUQUEN"
                Case "RNG"
                    Return "RIO NEGRO"
                Case "SCZ"
                    Return "SANTA CRUZ"
                Case "SDE"
                    Return "SANTIAGO DEL ESTERO"
                Case "SFE"
                    Return "SANTA FE"
                Case "SJN"
                    Return "SAN JUAN"
                Case "SLA"
                    Return "SALTA"
                Case "SLS"
                    Return "SAN LUIS"
                Case "TDF"
                    Return "TIERRA DEL FUEGO"
                Case "TUC"
                    Return "TUCUMAN"
                Case Else
                    Return ""
            End Select
        End Get

    End Property
    Public ReadOnly Property Pais() As Pais
        Get
            If _Pais Is Nothing Then
                _Pais = New Pais(cn, dr1("cry_0").ToString)
            End If

            Return _Pais
        End Get
    End Property
    Public Property CallejeroGCBA() As String
        Get
            Return dr1("bpagcba_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            dr1.BeginEdit()
            dr1("bpagcba_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public WriteOnly Property dirDefecto() As Integer
        Set(ByVal value As Integer)
            dr1.BeginEdit()
            dr1("BPAADDFLG_0") = value
            dr1.EndEdit()
        End Set
    End Property
    Public Property EntregaFactura() As Boolean
        Get
            Return CBool(IIf(CInt(dr1("XENTFAC_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr1.BeginEdit()
            dr1("XENTFAC_0") = IIf(value, 2, 1)
            dr1.EndEdit()
        End Set
    End Property
    Public Property AlturaGCBA() As Integer
        Get
            Return CInt(dr1("bpaaddnro_0"))
        End Get
        Set(ByVal value As Integer)
            dr1.BeginEdit()
            dr1("bpaaddnro_0") = value
            dr1.EndEdit()
        End Set
    End Property
    Public Property Telefono() As String
        Get
            Return dr1("tel_0").ToString
        End Get
        Set(ByVal value As String)
            dr1.BeginEdit()
            dr1("tel_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Fax() As String
        Get
            Return dr1("fax_0").ToString
        End Get
    End Property
    Public Property Mail() As String
        Get
            Return dr1("web_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("web_0") = IIf(value.Trim = "", " ", value.ToLower)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Portero() As String
        Get
            Return dr1("xportero_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("xportero_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Mail_portero() As String
        Get
            Return dr1("xpor_mail_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("xpor_mail_0") = IIf(value.Trim = "", " ", value.ToLower)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Identificador() As String
        Get
            Return dr1("extnum_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("extnum_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Telefono_Portero() As String
        Get
            Return dr1("xpor_tel_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("xpor_tel_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property Direccion_portero() As String
        Get
            Return dr1("xpor_add_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("xpor_add_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public Property MailFC() As String
        Get
            Return dr1("xmailfc_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            dr1.BeginEdit()
            dr1("xmailfc_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr1.EndEdit()
        End Set
    End Property
    Public ReadOnly Property TieneControlPeriodico() As Boolean
        Get
            Dim Sql As String = "SELECT * FROM xctrlt WHERE bpcnum_0 = :bpcnum_0 AND bpaadd_0 = :bpaadd_0"
            Dim da As New OracleDataAdapter(Sql, cn)
            Dim dt As New DataTable

            With da.SelectCommand.Parameters
                .Add("bpcnum_0", OracleType.VarChar).Value = Codigo
                .Add("bpaadd_0", OracleType.VarChar).Value = Sucursal
            End With

            da.Fill(dt)
            da.Dispose()

            Return dt.Rows.Count > 0

        End Get
    End Property
    Public ReadOnly Property CtrlCumpleIram() As Boolean
        Get
            Dim Sql As String
            Dim da As OracleDataAdapter
            Dim dt As DataTable
            Dim flg As Boolean = False

            'Consulta falta de carteles
            Sql = "SELECT * "
            Sql &= "FROM xctrls "
            Sql &= "WHERE (emergencia_0 > 0 OR baliza_0 > 0) AND bpcnum_0 = :bpcnum_0 AND bpaadd_0 = :bpaadd_0"
            dt = New DataTable
            da = New OracleDataAdapter(Sql, cn)
            With da.SelectCommand.Parameters
                .Add("bpcnum_0", OracleType.VarChar).Value = Codigo
                .Add("bpaadd_0", OracleType.VarChar).Value = Sucursal
            End With
            da.Fill(dt)
            da.Dispose()

            flg = dt.Rows.Count = 0

            If flg Then
                'Consulta falta de carteles
                Sql = "SELECT * "
                Sql &= "FROM xctrld "
                Sql &= "WHERE falta_0 = 'F' AND bpcnum_0 = :bpcnum_0 AND bpaadd_0 = :bpaadd_0"
                dt = New DataTable
                da = New OracleDataAdapter(Sql, cn)
                With da.SelectCommand.Parameters
                    .Add("bpcnum_0", OracleType.VarChar).Value = Codigo
                    .Add("bpaadd_0", OracleType.VarChar).Value = Sucursal
                End With

                da.Fill(dt)
                da.Dispose()

                flg = dt.Rows.Count = 0

            End If

            Return flg

        End Get
    End Property
    Public Property CumpleIram() As Boolean
        Get
            If CInt(dr1("iram_ok_0")) = 2 Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            dr1.BeginEdit()
            dr1("iram_ok_0") = IIf(value, 2, 1)
            dr1.EndEdit()
        End Set
    End Property
    Public ReadOnly Property SectorRuta() As Integer
        Get
            Dim dr As DataRow = dt2.Rows(0)
            Return CInt(dr("drn_0"))
        End Get
    End Property
    Public Property Transporte() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("extnum_0").ToString
            End If
            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("extnum_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property

    Public ReadOnly Property Ruta() As Integer
        Get
            Return CInt(dr2("drn_0"))
        End Get
    End Property
    Public Property TurnoMananaDesde() As String
        Get
            Dim txt As String = "0000"

            If dt2.Rows.Count > 0 Then
                txt = dr2("yhdesde1_0").ToString
            End If

            Return txt

        End Get
        Set(ByVal value As String)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("yhdesde1_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr2.EndEdit()
        End Set
    End Property
    Public Property TurnoMananaHasta() As String
        Get
            Dim txt As String = "0000"

            If dt2.Rows.Count > 0 Then
                txt = dr2("yhhasta1_0").ToString
            End If

            Return txt

        End Get
        Set(ByVal value As String)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("yhhasta1_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr2.EndEdit()
        End Set
    End Property
    Public Property TurnoTardeDesde() As String
        Get
            Dim txt As String = "0000"

            If dt2.Rows.Count > 0 Then
                txt = dr2("yhdesde2_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("yhdesde2_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr2.EndEdit()
        End Set
    End Property
    Public Property TurnoTardeHasta() As String
        Get
            Dim txt As String = "0000"

            If dt2.Rows.Count > 0 Then
                txt = dr2("yhhasta2_0").ToString
            End If

            Return txt

        End Get
        Set(ByVal value As String)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("yhhasta2_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr2.EndEdit()
        End Set
    End Property
    Public ReadOnly Property ModoEntrega() As String
        Get
            Dim x As String = dr2("mdl_0").ToString
            If x = " " Then x = "1"
            Return x
        End Get
    End Property
    Public ReadOnly Property PrioridadEntrega() As Integer
        Get
            Return CInt(dr2("dlvpio_0"))
        End Get
    End Property
    Public ReadOnly Property EsDireccionEntrega() As Boolean
        Get
            Return (dt2.Rows.Count = 1)
        End Get
    End Property
    Public Property SucursalPrincipal() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CBool(IIf(CInt(dr("bpaaddflg_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("bpaaddflg_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property SucursalEntregaActiva() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                flg = CBool(IIf(CInt(dr("enaflg_0")) = 2, True, False))
            End If

            Return flg

        End Get
    End Property

    Public Property AtencionLunes() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday1_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday1_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property AtencionMartes() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday2_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday2_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property AtencionMiercoles() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday3_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday3_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property AtencionJueves() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday4_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday4_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property AtencionViernes() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday5_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday5_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property AtencionSabado() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday6_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday6_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property AtencionDomingos() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("uvyday7_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("uvyday7_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property Expreso() As String
        Get
            If dr2 Is Nothing Then
                Return ""
            Else

                Return dr2("tranum_0").ToString
            End If

        End Get
        Set(ByVal value As String)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("tranum_0") = IIf(value.Trim = "", " ", value)
            dr2.EndEdit()
        End Set
    End Property
    Public Property Demora() As Integer
        Get
            If dr2 Is Nothing Then
                Return 0
            Else

                Return CInt(dr2("tmpesp_0"))
            End If
        End Get
        Set(ByVal value As Integer)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("tmpesp_0") = value
            dr2.EndEdit()
        End Set
    End Property
    Public Property BajaCliente() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("flgbaja_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("flgbaja_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property
    Public Property CantidadHidrantes() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("xhidrante_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("xhidrante_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property TipoSistema() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("xtiposist_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("xtiposist_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Hidrantes() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("xhidrante_0"))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property FechaInicioServicio() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("xiniservi_0"))
        End Get
    End Property
    Public Property SaleDesdeMunro() As Boolean
        Get
            If dr2 Is Nothing Then
                Return False
            Else
                Return CBool(IIf(CInt(dr2("xmunro_0")) <> 2, False, True))
            End If

        End Get
        Set(ByVal value As Boolean)
            If dr2 Is Nothing Then NuevaSucursalEntrega()

            dr2.BeginEdit()
            dr2("xmunro_0") = IIf(value, 2, 1)
            dr2.EndEdit()
        End Set
    End Property

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Liberar otro estado (objetos administrados).
            End If

            ' TODO: Liberar su propio estado (objetos no administrados).
            ' TODO: Establecer campos grandes como Null.
            da1.Dispose()
            dt1.Dispose()
        End If
        Me.disposedValue = True
    End Sub

    ' Visual Basic agregó este código para implementar correctamente el modelo descartable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' No cambie este código. Coloque el código de limpieza en Dispose (ByVal que se dispone como Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Dispose(False)
    End Sub

End Class