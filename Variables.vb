
Imports System.Data.SqlClient
Module Variables
    Public nivel As Integer
    Public sql, cuenta, empresa, codcli, nomcli, rtn, vcodigo, bodega, nbodega, tipo, vsucu, vnsucu, vproeli, vnproducto, vcompra, vpedido As String
    Public sql2, v_nombre_respaldo, pasador, pasador3, pasador2, pasadorh, pasadorl, pncredito, pndebito, letrasn, pcheques, ppartida, sql3, vfactura As String
    Public control, existencia, comproinv, diferencia, afacturar, existencia2, vproveedor As Integer
    Public vprecio, vsubtotalp, vsubototalf, impuestop, totalp, totalie, totali1, totali2, totali, totalf, vpasador, rebajaf, pcosto, vimpuestocosto, vcosto As Double
    Public almacen As DataSet
    Public ncant() As Integer
    Public Factu As Integer
    Public tipofactu As String
    Public nbodegas() As String
    Public ncod() As String
    Public almacen2 As DataSet
    Public resultado, vcontrold, vfila, alarmador, resta, diferencia2, vfila2, vfila3, veri, res As Integer
    Public resultado2, limite As Integer
    Public dt As DataTable
    Public dtrow, redistro As DataRow
    Public dtrow2 As DataRow
    Public vnivel, estadot, vnproveedor As String
    Public vnombre As String
    Public vusuario As String
    Public vfechanac, tfactura, tfecha, descpro, vcuenta, vcuenta2, vcuenta3, vcuenta4, vcuenta5, vcuenta6, vsalidabodega, pfactura, porden, vturno, pasador0, pasador1, ncuenta As String
    Public vtcompras As Double
    Public ds As DataSet
    Public bd As String
    Public estado, lleva, conteo As Integer
    Public row As DataRow
    Public cnn, cnn2 As New SqlConnection
    Public idusuario As Integer
    Public DA As SqlDataAdapter
    Public Query, cadena As String
    Public cmdsql As SqlCommand
    Sub alternarcolorfilardatagridview(ByVal dgv As DataGridView)
        Try
            With dgv
                .RowsDefaultCellStyle.BackColor = Color.White
                .AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue
            End With
        Catch ex As Exception

        End Try
    End Sub
    Sub conectado()
        Try
            If empresa = "Traconsa" Then
                cnn = New SqlConnection("Data Source=192.168.1.10;Initial Catalog=traconsa2016;User ID=sistemas;Password=S1stemas2016")
                cnn.Open()
                bd = "traconsa2016"
                cadena = "Data Source=192.168.1.10;Initial Catalog=traconsa2016;User ID=sistemas;Password=S1stemas2016"
            ElseIf empresa = "Ecresa" Then
                MsgBox("Ecresa")
                cnn = New SqlConnection("Data Source=192.168.1.10;Initial Catalog=Ecresa2016;User ID=sistemas;Password=S1stemas2016")
                cnn.Open()
                bd = "Ecresa2016"
                cadena = "Data Source=192.168.1.10;Initial Catalog=Ecresa2016;User ID=sistemas;Password=S1stemas2016"
            Else
                cnn = New SqlConnection("Data Source=192.168.1.10;Initial Catalog=fares_pruebas;User ID=sistemas;Password=S1stemas2016")
                cnn.Open()
                bd = "fares_pruebas"
                cadena = "Data Source=192.168.1.10;Initial Catalog=fares_pruebas;User ID=sistemas;Password=S1stemas2016"
            End If


        Catch ex As Exception
            MsgBox(ex.Message)

        End Try


    End Sub


    Sub desconectado()
        Try

            If cnn.State = ConnectionState.Open Then
                    cnn.Close()

                Else

                End If


        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub



End Module
