Public Shared Function Reserva_Stock_From_MI3(ByRef pStockResSolicitud As clsBeStock_res,
                                                  ByVal DiasVencimiento As Double,
                                                  ByVal MaquinaQueSolicita As String,
                                                  ByVal pBeConfigEnc As clsBeI_nav_config_enc,
                                                  ByRef pCantidadDisponibleStock As Double,
                                                  ByVal pIdPropietarioBodega As Integer,
                                                  ByRef pListStockResOUT As List(Of clsBeStock_res),
                                                  ByRef lConnection As SqlConnection,
                                                  ByRef ltransaction As SqlTransaction,
                                                  Optional No_Linea As Integer = 0,
                                                  Optional pTarea_Reabasto As Boolean = False,
                                                  Optional ByVal pBeTrasladoDet As clsBeI_nav_ped_traslado_det = Nothing) As Boolean

        Reserva_Stock_From_MI3 = False

        Try

#Region "Variables"

            Dim lBeStockExistente As New List(Of clsBeStock)
            Dim lBeStockExistenteZonasNoPicking As New List(Of clsBeStock)
            Dim lBeStockExistenteZonaPicking As New List(Of clsBeStock)
            Dim lBeStockExistenteZonaPickingPresentacion As New List(Of clsBeStock)
            Dim lBeStockExistenteTmp As New List(Of clsBeStock)
            Dim lBeStockExistenteTomeDesde As New List(Of clsBeStock)
            Dim lBeStockDisponible As New List(Of clsBeStock)
            Dim lBeStockAReservar As New List(Of clsBeStock_res)
            Dim vIndicePresentacion As Integer = -1
            Dim vIndiceUbicacion As Integer = 0
            Dim vCantidadReservada As Double = 0
            Dim vPesoReservado As Double = 0
            Dim BeProducto As New clsBeProducto
            Dim BeStockRes As New clsBeStock_res
            Dim vCantidadCompletada As Boolean = False
            Dim vCantidadDispStock As Double = 0
            Dim vCantidadPendiente As Double = 0
            Dim vCantidadSolicitadaPedido As Double = 0
            Dim vValorEnteroCantidadSolicitadaPedido As Integer = 0
            Dim vCantidadAReservarPorIdStock As Double = 0
            Dim vCantidadEnteraPres As Integer = 0
            Dim vCantidadDecimalUMBas As Double = 0
            Dim vCantidadStock As Double = 0
            Dim vCantidadStockEnPres As Double = 0
            Dim vDisponibleStockEnPres As Double = 0
            Dim vPesoStock As Double = 0
            Dim vPesoPendiente As Double = 0
            Dim vPesoSolicitadoPedido As Double = 0
            Dim vPesoAReservarPorIdStock As Double = 0
            Dim Idx As Integer = 0
            Dim vCantidadEnStockEnPres As Double = 0
            Dim vCantidadSolicitadaPedidoEnPres As Double = 0
            Dim vCantidadEnteraStockPres As Double = 0
            Dim vCantidadDecimalStockUMBas As Double = 0
            Dim vPesoEnteroPres As Double = 0
            Dim vPesoDecimalUMBas As Double = 0
            Dim vPesoEnteroPresStock As Double = 0
            Dim vPesoDecimalStockUMBas As Double = 0
            Dim vCantidadDecimalUMBasStock As Double = 0
            Dim BeStockDestino As New clsBeStock
            Dim BeUbicacionStock As New clsBeBodega_ubicacion()
            Dim vCantidadPendienteEnPres As Double = 0
            Dim vCantidadDispStockEnPres As Double = 0
            Dim CantidadEnUMBasPorPresentacionDelStock As Double = 0
            Dim vCantidadEnteraSolicitadaPedidoEnPres As Double = 0
            Dim vCantidadDecimalSolicitadaPedidoEnPres As Double = 0
            Dim IdProducto As Integer = 0
            Dim vSolicitudEsEnUMBas As Boolean = False
            Dim BeStockOriginal As New clsBeStock()
            Dim vOrdernarListaStockSinPresentacionPrimero As Boolean = False
            Dim vConvirtioCantidadSolicitadaEnUmBas As Boolean = False
            Dim vlBeStockAReservarUMBas As New List(Of clsBeStock_res)
            Dim vlBeStockAReservarPresFaltante As New List(Of clsBeStock_res)
            Dim vCantidadTarimasCompletasAPickearClavaud As Double = 0
            Dim vCantidadEnteraTarimasCompletasClavaud As Double = 0
            Dim vCantidadDecimalTarimasCompletasClavaud As Double = 0
            Dim BeUbicacionEnMemoria As New clsBeBodega_ubicacion()
            Dim vCantidadProductoPorTarima As Double = 0
            Dim vResultCalculoTarimaEstaCompleta As Double = 0
            Dim vCantidadEnStockEnPresentacionClavaud As Double = 0
            Dim lBeStockConPalletsCompletosClavaud As New List(Of clsBeStock)
            Dim lBeStockConPalletsInCompletosClavaud As New List(Of clsBeStock)
            Dim lBeStockZonaPicking As New List(Of clsBeStock)
            Dim lBeStockZonasNoPicking As New List(Of clsBeStock)
            Dim vCantPVConPalletsCompletosClavaud As Integer = 0 'Próximos a vencer
            Dim vCantPVConPalletsInCompletosClavaud As Integer = 0 'Próximos a vencer
            Dim vBusquedaEnUmBas As Boolean = False
            Dim vZonaNoPickingStockEnUmBas As Boolean = False
            Dim vRefCantidadReservada As Double = 0
            Dim vRestoInventarioEnUmBas As Boolean = False
            Dim BePresentacionDefecto As New clsBeProducto_Presentacion
            Dim vEncontroExistenciaEnPresentacion As Boolean = False
            Dim CantidadStockDestino As Double = 0
            Dim BePedidoDet As New clsBeTrans_pe_det
            Dim vFechaDefecto As Date = New Date(1900, 1, 1)
            Dim FechaMinimaVenceStock As Date = vFechaDefecto
            Dim ExcepcionFechaVenceEsInferiorEnZonaPicking As Boolean = False
            Dim vFechaMinimaVenceZonaPicking As Date = vFechaDefecto
            Dim vFechaMinimaVenceZonaALM As Date = vFechaDefecto
            Dim vVenceMinimaPickingCompletoClavaud As Date = vFechaDefecto
            Dim vVenceMinimaPickingInCompletoClavaud As Date = vFechaDefecto
            Dim ListaEstadosDeProceso As New List(Of Integer)
            Dim vPermitirDecimales As Boolean = False
            Dim BeBodega As New clsBeBodega
            Dim vCantidadStockZonaNoPicking As Double = 0
            Dim vCantidadStockZonaPicking As Double = 0
            Dim vMensajeNoExplosionEnZonasNoPicking = ""
            Dim vMensajeReserva = ""
            Dim vCantidadTotalStock As Double = 0
            Dim vStockDispZonaPicking As Integer = 0
            Dim vFechaMinima As Date = vFechaDefecto
            Dim Iniciar_En As Integer = 0
            Dim pStockResBusquedaParaExplosion As New clsBeStock_res
            Dim vRestoStockReservado As Boolean = False
            Dim vProcessResult As New List(Of String)
            Dim vIdTipoPedido As Integer = 0
            Dim pEs_Devolucion As Boolean = False
            Dim vPresReserva As Integer = 0

#End Region
            vIdTipoPedido = clsLnTrans_pe_enc.Get_IdTipoPedido_By_IdPedidoEnc(pStockResSolicitud.IdPedido,
                                                                              lConnection,
                                                                              ltransaction)

            pEs_Devolucion = (vIdTipoPedido = clsDataContractDI.tTipoDocumentoSalida.Devolucion_Proveedor)

            vPresReserva = pStockResSolicitud.IdPresentacion

            Cargar_Bodega_Y_Linea_Pedido(pBeConfigEnc.Idbodega,
                                         pIdPropietarioBodega,
                                         pStockResSolicitud.IdPedido,
                                         pStockResSolicitud.IdPedidoDet,
                                         lConnection,
                                         ltransaction,
                                         BeBodega,
                                         BePedidoDet)

            Get_Objetos_Producto(pStockResSolicitud,
                                 BePresentacionDefecto,
                                 IdProducto,
                                 BeProducto,
                                 lConnection,
                                 ltransaction)

            Dim ListasStock = Obtener_Listas_De_Stock(pStockResSolicitud,
                                                      BeProducto,
                                                      DiasVencimiento,
                                                      pBeConfigEnc,
                                                      lConnection,
                                                      ltransaction,
                                                      pTarea_Reabasto,
                                                      pEs_Devolucion)

            lBeStockExistente = ListasStock.lBeStockExistente
            lBeStockExistenteZonasNoPicking = ListasStock.lBeStockExistenteZonasNoPicking
            lBeStockExistenteZonaPicking = ListasStock.lBeStockExistenteZonaPicking

            If pStockResSolicitud.IdPresentacion <> 0 AndAlso lBeStockExistenteZonaPicking IsNot Nothing AndAlso lBeStockExistenteZonaPicking.Count > 0 Then
                lBeStockExistenteZonaPicking = lBeStockExistenteZonaPicking.FindAll(Function(x) x.UbicacionPicking = True)
            End If

            vFechaMinimaVenceZonaPicking = vFechaDefecto
            vFechaMinimaVenceZonaALM = vFechaDefecto
            vVenceMinimaPickingCompletoClavaud = vFechaDefecto
            vVenceMinimaPickingInCompletoClavaud = vFechaDefecto

INICIAR_CON_NUEVO_LSTOCK:

            If Not lBeStockExistente Is Nothing Then

#Region "RESTAR_STOCK_RESERVADO"

                Procesar_Y_Restar_Stock_Reservado(lBeStockExistente,
                                                  lPresentaciones,
                                                  vEncontroExistenciaEnPresentacion,
                                                  vCantidadProductoPorTarima,
                                                  vCantidadTarimasCompletasAPickearClavaud,
                                                  vCantidadEnteraTarimasCompletasClavaud,
                                                  vCantidadDecimalTarimasCompletasClavaud,
                                                  pStockResSolicitud,
                                                  vOrdernarListaStockSinPresentacionPrimero,
                                                  pBeConfigEnc,
                                                  lConnection,
                                                  ltransaction)

                Procesar_Y_Restar_Stock_Reservado(lBeStockExistenteZonasNoPicking,
                                                  lPresentaciones,
                                                  vEncontroExistenciaEnPresentacion,
                                                  vCantidadProductoPorTarima,
                                                  vCantidadTarimasCompletasAPickearClavaud,
                                                  vCantidadEnteraTarimasCompletasClavaud,
                                                  vCantidadDecimalTarimasCompletasClavaud,
                                                  pStockResSolicitud,
                                                  vOrdernarListaStockSinPresentacionPrimero,
                                                  pBeConfigEnc,
                                                  lConnection,
                                                  ltransaction)

                Procesar_Y_Restar_Stock_Reservado(lBeStockExistenteZonaPicking,
                                                  lPresentaciones,
                                                  vEncontroExistenciaEnPresentacion,
                                                  vCantidadProductoPorTarima,
                                                  vCantidadTarimasCompletasAPickearClavaud,
                                                  vCantidadEnteraTarimasCompletasClavaud,
                                                  vCantidadDecimalTarimasCompletasClavaud,
                                                  pStockResSolicitud,
                                                  vOrdernarListaStockSinPresentacionPrimero,
                                                  pBeConfigEnc,
                                                  lConnection,
                                                  ltransaction)

                vRestoStockReservado = True

#End Region

#Region "OBTENER_FECHA_MINIMA_DE_INVENTARIO"

                FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                 DiasVencimiento,
                                                                                 pBeConfigEnc,
                                                                                 lConnection,
                                                                                 ltransaction,
                                                                                 BeProducto,
                                                                                 pTarea_Reabasto,
                                                                                 vFechaMinimaVenceZonaPicking,
                                                                                 vFechaMinimaVenceZonaALM,
                                                                                 lBeStockExistente,
                                                                                 BePresentacionDefecto)

                If lBeStockExistenteZonasNoPicking IsNot Nothing AndAlso lBeStockExistenteZonasNoPicking.Count > 0 AndAlso vFechaMinimaVenceZonaALM < FechaMinimaVenceStock Then
                    lBeStockExistente = lBeStockExistenteZonasNoPicking
                End If

                If FechaMinimaVenceStock.Date = vFechaDefecto AndAlso lBeStockExistente IsNot Nothing AndAlso lBeStockExistente.Count > 0 Then
                    FechaMinimaVenceStock = lBeStockExistente.Min(Function(x) x.Fecha_vence)
                End If
#End Region


EXPLOSIONAR_PRODUCTO:

                If Stock_Requiere_Explosion(pBeConfigEnc, lBeStockExistente, pStockResSolicitud) Then

                    If pStockResSolicitud.IdPresentacion = 0 Then
                        vOrdernarListaStockSinPresentacionPrimero = True : vBusquedaEnUmBas = True
                    End If

                    pStockResBusquedaParaExplosion = Nothing

                    If lBeStockExistente.Count = 0 AndAlso lBeStockExistenteZonaPicking.Count = 0 AndAlso lBeStockExistenteZonasNoPicking.Count > 0 Then
                        lBeStockExistente = lBeStockExistenteZonasNoPicking
                    End If

                    If lBeStockExistente.Count = 0 Then
                        If pStockResSolicitud.IdPresentacion = 0 Then
                            If BePresentacionDefecto IsNot Nothing Then
                                pStockResBusquedaParaExplosion = pStockResSolicitud.Clone()
                                pStockResBusquedaParaExplosion.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                vBusquedaEnUmBas = False
                            Else
                                '#CKFK20240320 no aplica lanzar excepcion por esto
                                ' Throw New Exception("ERROR_202302021127: Se está intentando reservar inventario que requiere explosión pero no está definida la presentación por defecto del producto: " & BeProducto.Codigo)
                            End If
                        ElseIf Not vEncontroExistenciaEnPresentacion Then
                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                Throw New Exception(String.Format("Error_202212140140D: {0} Código: {1} Sol: {2} Disp: {3}. ", clsDalEx.ErrorS0002, BeProducto.Codigo, pStockResSolicitud.Cantidad, 0))
                            ElseIf Not vCantidadCompletada Then
                                '#EJC202312141100:Agregado por BYB.
                                If BePedidoDet.IdPresentacion = 0 Then
                                    vBusquedaEnUmBas = True
                                End If
                            End If
                        Else
                            vBusquedaEnUmBas = True
                        End If
                    ElseIf Not vCantidadCompletada AndAlso (pStockResSolicitud.IdPresentacion = 0 OrElse (BePedidoDet IsNot Nothing AndAlso BePedidoDet.IdPresentacion = 0)) Then
                        vBusquedaEnUmBas = True
                    End If

                    If vBusquedaEnUmBas Then
                        ' Dividir la cantidad solicitada en su parte entera y decimal
                        Split_Decimal(pStockResSolicitud.Cantidad, vCantidadEnteraPres, vCantidadDecimalUMBas)

                        '#CKFK20240126 Agregué esta validación porque si no tiene presentacion da error de object not reference
                        If BePresentacionDefecto IsNot Nothing Then
                            ' Ajustar la cantidad decimal según el factor de presentación y redondear hacia arriba
                            vCantidadDecimalUMBas = Math.Ceiling(Math.Round(vCantidadDecimalUMBas * BePresentacionDefecto.Factor, 2))
                        End If

                        If vCantidadEnteraPres > 0 AndAlso pStockResSolicitud.IdPresentacion <> 0 Then
                            ' Calcular la cantidad solicitada basada en la presentación y su factor
                            vCantidadSolicitadaPedido = vCantidadEnteraPres * BePresentacionDefecto.Factor

                            ' Si no se encontró existencia en la presentación, añadir la cantidad decimal en UMBas a la cantidad solicitada
                            If Not vEncontroExistenciaEnPresentacion Then
                                vCantidadSolicitadaPedido += vCantidadDecimalUMBas
                                vCantidadDecimalUMBas = 0
                            End If
                        ElseIf vCantidadEnteraPres > 0 Then
                            ' Si no hay presentación específica y hay cantidad entera, usar la cantidad total solicitada
                            vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad
                        Else
                            ' Si no hay una cantidad entera, usar la cantidad decimal en UMBas
                            vCantidadSolicitadaPedido = vCantidadDecimalUMBas
                        End If

                        ' Actualizar la solicitud de stock con la cantidad calculada
                        pStockResSolicitud.Cantidad = vCantidadSolicitadaPedido
                        pStockResSolicitud.Atributo_Variante_1 = Nothing
                        pStockResSolicitud.IdPresentacion = 0
                    End If

                    If lBeStockExistente.Count = 0 Then

                        If Not vBusquedaEnUmBas AndAlso pStockResBusquedaParaExplosion Is Nothing Then
                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                Throw New Exception(String.Format("Error_202212140140D: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                           BeProducto.Codigo,
                                                                           pStockResSolicitud.Cantidad,
                                                                           0))
                            Else
                                If Not vCantidadCompletada Then
                                    Dim vMensajeError20230306 As String = String.Format("Error202303051226: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                    BeProducto.Codigo,
                                                                                    vCantidadSolicitadaPedido,
                                                                                    vCantidadStock)
                                    clsLnLog_error_wms.Agregar_Error(vMensajeError20230306 & "C se realizó exit function con Reserva_Stock_From_MI3 = false", lConnection, ltransaction)
                                    'Exit Function
                                End If
                            End If
                        End If

                        '#EJC202312191315: Para BYB, con amor.
                        If Not vBusquedaEnUmBas AndAlso pStockResBusquedaParaExplosion Is Nothing Then

                            If lBeStockExistente.Count = 0 Then
                                If BePresentacionDefecto IsNot Nothing Then
                                    pStockResBusquedaParaExplosion = pStockResSolicitud.Clone()
                                    pStockResBusquedaParaExplosion.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                    vBusquedaEnUmBas = False
                                Else
                                    '#CKFK20240320 no aplica lanzar excepcion por esto
                                    'Throw New Exception("ERROR_202302021127: Se está intentando reservar inventario que requiere explosión pero no está definida la presentación por defecto del producto: " & BeProducto.Codigo)
                                End If
                            End If

                        End If

                        '#EJC202309271639: Se busca explosionar primero de zonas de picking.
                        lBeStockExistenteZonaPicking = clsLnStock.lStock(IIf(vBusquedaEnUmBas,
                                                                  pStockResSolicitud,
                                                                  pStockResBusquedaParaExplosion),
                                                              BeProducto,
                                                              DiasVencimiento,
                                                              pBeConfigEnc,
                                                              lConnection,
                                                              ltransaction,
                                                              False,
                                                              True,
                                                              pTarea_Reabasto,
                                                              pEs_Devolucion)

                        Restar_Stock_Reservado(lBeStockExistenteZonaPicking,
                                               pBeConfigEnc,
                                               lConnection,
                                               ltransaction)

                        If lBeStockExistenteZonaPicking IsNot Nothing AndAlso lBeStockExistenteZonaPicking.Any() Then
                            vFechaMinimaVenceZonaPicking = lBeStockExistenteZonaPicking.Min(Function(x) x.Fecha_vence)
                            vProcessResult.Add("#MI3_2312201855: Se encontraron " & lBeStockExistenteZonaPicking.Count & " registros. La fecha mínima de picking es: " & vFechaMinimaVenceZonaPicking)
                        End If

                        lBeStockExistenteZonasNoPicking = clsLnStock.lStock(IIf(vBusquedaEnUmBas,
                                                                  pStockResSolicitud,
                                                                  pStockResBusquedaParaExplosion),
                                                              BeProducto,
                                                              DiasVencimiento,
                                                              pBeConfigEnc,
                                                              lConnection,
                                                              ltransaction,
                                                              True,
                                                              True,
                                                              pTarea_Reabasto,
                                                              pEs_Devolucion)

                        Restar_Stock_Reservado(lBeStockExistenteZonasNoPicking,
                                               pBeConfigEnc,
                                               lConnection,
                                               ltransaction)

                        If lBeStockExistenteZonasNoPicking IsNot Nothing AndAlso lBeStockExistenteZonasNoPicking.Any() Then
                            vFechaMinimaVenceZonaALM = lBeStockExistenteZonasNoPicking.Min(Function(x) x.Fecha_vence)
                        End If

                        vRestoInventarioEnUmBas = True

                    End If

                    If vFechaMinimaVenceZonaALM > vFechaMinimaVenceZonaPicking Then
                        If lBeStockExistenteZonaPicking IsNot Nothing AndAlso lBeStockExistenteZonaPicking.Any() Then
                            lBeStockExistente = lBeStockExistenteZonaPicking
                        End If
                    Else
                        If lBeStockExistenteZonasNoPicking IsNot Nothing AndAlso lBeStockExistenteZonasNoPicking.Any() Then
                            lBeStockExistente = lBeStockExistenteZonasNoPicking
                            Iniciar_En = 3
                        End If
                    End If

                    'Por aquí voy en la factorización- 202312061325

                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                    vCantidadStockZonaNoPicking = lBeStockExistenteZonasNoPicking.Sum(Function(x) x.Cantidad)
                    vCantidadStockZonaPicking = lBeStockExistenteZonaPicking.Sum(Function(x) x.Cantidad)
                    vCantidadTotalStock = vCantidadStockZonaNoPicking + vCantidadStockZonaPicking

                    If pStockResSolicitud.IdPresentacion = 0 AndAlso vCantidadSolicitadaPedido = 0 Then
                        vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad
                    End If

                    If (vBusquedaEnUmBas AndAlso lBeStockExistente.Count = 0) AndAlso (vCantidadSolicitadaPedido > vCantidadTotalStock) AndAlso (pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si) Then

                        Throw New Exception(String.Format("Error_202212140140D: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                       BeProducto.Codigo,
                                                                       pStockResSolicitud.Cantidad,
                                                                       0))
                    Else

                        If lBeStockExistente.Count = 0 Then

                            If pStockResSolicitud.IdPresentacion = 0 Then vBusquedaEnUmBas = True

                            '#CKFK20231009 Puse el conmutar en false porque aqui solo quiero unidades
                            lBeStockExistente = clsLnStock.lStock(pStockResSolicitud,
                                                                  BeProducto,
                                                                  DiasVencimiento,
                                                                  pBeConfigEnc,
                                                                  lConnection,
                                                                  ltransaction,
                                                                  True,
                                                                  False,
                                                                  pTarea_Reabasto,
                                                                  pEs_Devolucion)

                            Restar_Stock_Reservado(lBeStockExistente,
                                                   pBeConfigEnc,
                                                   lConnection,
                                                   ltransaction)

                            vRestoInventarioEnUmBas = True

                            lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                            If lBeStockExistente.Count > 0 Then
                                '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3
                                FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                 DiasVencimiento,
                                                                                                 pBeConfigEnc,
                                                                                                 lConnection,
                                                                                                 ltransaction,
                                                                                                 BeProducto,
                                                                                                 pTarea_Reabasto,
                                                                                                 vFechaMinimaVenceZonaPicking,
                                                                                                 vFechaMinimaVenceZonaALM,
                                                                                                 lBeStockExistente,
                                                                                                 BePresentacionDefecto)
                            Else
                                If Not lBeStockExistenteZonaPicking.Count = 0 Then
                                    lBeStockExistente = lBeStockExistenteZonaPicking
                                End If
                            End If

                            vZonaNoPickingStockEnUmBas = lBeStockExistente.Count > 0

                        End If

                        If lBeStockExistente.Count = 0 Then

                            If pStockResSolicitud.IdPresentacion = 0 Then vBusquedaEnUmBas = True

                            '#EJC202309121412: Buscar en zonas de picking  unidades
                            lBeStockExistente = clsLnStock.lStock(pStockResSolicitud,
                                                                  BeProducto,
                                                                  DiasVencimiento,
                                                                  pBeConfigEnc,
                                                                  lConnection,
                                                                  ltransaction,
                                                                  False,
                                                                  True,
                                                                  pTarea_Reabasto,
                                                                  pEs_Devolucion)

                            Restar_Stock_Reservado(lBeStockExistente,
                                                   pBeConfigEnc,
                                                   lConnection,
                                                   ltransaction)

                            vRestoInventarioEnUmBas = True

                            lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                            FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                             DiasVencimiento,
                                                                                             pBeConfigEnc,
                                                                                             lConnection,
                                                                                             ltransaction,
                                                                                             BeProducto,
                                                                                             pTarea_Reabasto,
                                                                                             vFechaMinimaVenceZonaPicking,
                                                                                             vFechaMinimaVenceZonaALM,
                                                                                             lBeStockExistente,
                                                                                             BePresentacionDefecto)

                        End If

                        If vBusquedaEnUmBas AndAlso lBeStockExistente.Count = 0 Then

                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                Throw New Exception(String.Format("Error_202212140140D: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                       BeProducto.Codigo,
                                                                       pStockResSolicitud.Cantidad,
                                                                       0))
                            Else

                                If Not vCantidadCompletada Then

                                    If Not ListaEstadosDeProceso.Contains(105) Then

                                        If (pBeConfigEnc.Explosion_Automatica) AndAlso (lBeStockExistente.Count = 0 OrElse pStockResSolicitud.IdPresentacion = 0) Then

                                            If Not BePresentacionDefecto Is Nothing Then '#EJC20230202: Entonces voy a buscar inventario en cajas (con la presentación por defecto si existe)
                                                clsPublic.CopyObject(pStockResSolicitud, pStockResBusquedaParaExplosion)
                                                pStockResBusquedaParaExplosion = pStockResSolicitud.Clone()
                                                pStockResBusquedaParaExplosion.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                vBusquedaEnUmBas = False
                                            Else
                                                '#CKFK20240320 no aplica lanzar excepcion por esto
                                                'Throw New Exception("ERROR_202302021127: Se está intentando reservar inventario que requiere explosión pero no está definida la presentación por defecto del producto: " & BeProducto.Codigo)
                                            End If

                                            lBeStockExistente = clsLnStock.lStock(pStockResBusquedaParaExplosion,
                                                                                  BeProducto,
                                                                                  DiasVencimiento,
                                                                                  pBeConfigEnc,
                                                                                  lConnection,
                                                                                  ltransaction,
                                                                                  True,
                                                                                  True,
                                                                                  pTarea_Reabasto,
                                                                                  pEs_Devolucion)

                                            Restar_Stock_Reservado(lBeStockExistente,
                                                                    pBeConfigEnc,
                                                                    lConnection,
                                                                    ltransaction)

                                            vRestoInventarioEnUmBas = True

                                            lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                            FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                             DiasVencimiento,
                                                                                                             pBeConfigEnc,
                                                                                                             lConnection,
                                                                                                             ltransaction,
                                                                                                             BeProducto,
                                                                                                             pTarea_Reabasto,
                                                                                                             vFechaMinimaVenceZonaPicking,
                                                                                                             vFechaMinimaVenceZonaALM,
                                                                                                             lBeStockExistente,
                                                                                                             BePresentacionDefecto)

                                            ListaEstadosDeProceso.Add(105)

                                        Else

                                            ListaEstadosDeProceso.Add(105)

                                            Dim vMensajeError20230306 As String = String.Format("Error202303051227: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                        BeProducto.Codigo,
                                                                                        vCantidadSolicitadaPedido,
                                                                                        vCantidadStock)
                                            clsLnLog_error_wms.Agregar_Error(vMensajeError20230306 & "D se realizó exit function con Reserva_Stock_From_MI3 = false")
                                            Exit Function

                                        End If

                                    Else

                                        Dim vMensajeError20230306 As String = String.Format("Error202303051227: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                        BeProducto.Codigo,
                                                                                        vCantidadSolicitadaPedido,
                                                                                        vCantidadStock)
                                        clsLnLog_error_wms.Agregar_Error(vMensajeError20230306 & "D se realizó exit function con Reserva_Stock_From_MI3 = false")
                                        Exit Function

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

                If BeProducto.Codigo = "WMS223" Then
                    Debug.Write("Espera")
                End If

                If lBeStockExistente.Count > 0 Then

                    If pStockResSolicitud.IdPresentacion = 0 Then
                        vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad
                    Else

                        BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = pStockResSolicitud.IdPresentacion}

                        vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                        If vIndicePresentacion <> -1 Then
                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                            BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                        Else
                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                            BePresentacionDefecto.IdPresentacion = pStockResSolicitud.IdPresentacion
                            clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                            lPresentaciones.Add(BePresentacionDefecto.Clone())
                        End If

                        If BePresentacionDefecto.EsPallet Then

                            Dim vFactorPallet As Double = (BePresentacionDefecto.Factor * BePresentacionDefecto.CajasPorCama * BePresentacionDefecto.CamasPorTarima)

                            If vFactorPallet > 0 Then
                                vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad * vFactorPallet
                            Else
                                Throw New Exception("No se pudo reservar el stock para el tipo de producto pallet porque los factores de conversión dan un denominador = 0")
                            End If

                        Else

                            If (pBeConfigEnc.Explosion_Automatica) Then

                                Split_Decimal(pStockResSolicitud.Cantidad,
                                              vCantidadEnteraPres,
                                              vCantidadDecimalUMBas)


                                vCantidadDecimalUMBas = Math.Ceiling(Math.Round(vCantidadDecimalUMBas * BePresentacionDefecto.Factor, 2))

                                If vCantidadEnteraPres > 0 Then
                                    vCantidadSolicitadaPedido = vCantidadEnteraPres
                                Else
                                    vCantidadSolicitadaPedido = vCantidadDecimalUMBas
                                End If

                            Else
                                vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad * BePresentacionDefecto.Factor
                                '#EJC20240808: Asignar cantidad pendiente también en base a cantidad solicitada de pedido en umbas.
                                vCantidadPendiente = vCantidadSolicitadaPedido
                                vConvirtioCantidadSolicitadaEnUmBas = True
                            End If

                            If Not clsLnProducto.Tiene_Control_Por_Peso_By_IdProductoBodega(pStockResSolicitud.IdProductoBodega,
                                                                                            lConnection,
                                                                                            ltransaction) Then

                                If Integer.TryParse(vCantidadSolicitadaPedido, vValorEnteroCantidadSolicitadaPedido) Then
                                    vCantidadSolicitadaPedido = vValorEnteroCantidadSolicitadaPedido
                                Else
                                    vCantidadSolicitadaPedido = Math.Truncate(vCantidadSolicitadaPedido)
                                End If

                            Else
                                vCantidadSolicitadaPedido = Math.Round(vCantidadSolicitadaPedido, 6)
                            End If

                        End If

                    End If

                    If Not lBeStockExistente Is Nothing Then

                        vCantidadStock = lBeStockExistente.Sum(Function(x) x.Cantidad)
                        vDisponibleStockEnPres = lBeStockExistente.Where(Function(x) x.IdPresentacion <> 0).Sum(Function(x) x.Cantidad)
                        vPesoStock = lBeStockExistente.Sum(Function(x) x.Peso)

                        '#CKFK20250603 Agregué esta condición para que encuentre el inventario de moldes
                        vEncontroExistenciaEnPresentacion = vDisponibleStockEnPres > 0

                        If pBeConfigEnc.Explosion_Automatica Then

                            If pStockResSolicitud.IdPresentacion = 0 Then

                                vCantidadStockZonaNoPicking = lBeStockExistenteZonasNoPicking.Sum(Function(x) x.Cantidad)
                                vCantidadStockZonaPicking = lBeStockExistenteZonaPicking.Sum(Function(x) x.Cantidad)
                                vCantidadTotalStock = vCantidadStockZonaNoPicking + vCantidadStockZonaPicking

                                If (vCantidadSolicitadaPedido > vCantidadStock) AndAlso (vCantidadSolicitadaPedido > vCantidadTotalStock) Then

                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                    BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                              lConnection,
                                                                                                                              ltransaction)

                                    If BePresentacionDefecto Is Nothing Then

                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                            Throw New Exception(String.Format("Error_202250607:  {0} Código:  {1} Sol: {2} Disp: {3}. " & vbNewLine,
                                                                              "El producto no tiene presentación, no se puede explosionar",
                                                                               BeProducto.Codigo,
                                                                               vCantidadSolicitadaPedido,
                                                                               vCantidadTotalStock))
                                        End If

                                    End If

                                    Dim BeStockResUMBas As New clsBeStock_res
                                    BeStockResUMBas = pStockResSolicitud.Clone()
                                    BeStockResUMBas.IdPresentacion = BePresentacionDefecto.IdPresentacion

                                    lBeStockExistenteTmp = clsLnStock.lStock(BeStockResUMBas,
                                                                             BeProducto,
                                                                             DiasVencimiento,
                                                                             pBeConfigEnc,
                                                                             lConnection,
                                                                             ltransaction,
                                                                             IIf(pTarea_Reabasto, True, False),
                                                                             True,
                                                                             pTarea_Reabasto,
                                                                             pEs_Devolucion)

                                    Restar_Stock_Reservado(lBeStockExistenteTmp,
                                                           pBeConfigEnc,
                                                           lConnection,
                                                           ltransaction)


                                    lBeStockExistenteTmp = lBeStockExistenteTmp.FindAll(Function(x) x.Cantidad > 0)

                                    vCantidadStock = lBeStockExistenteTmp.Sum(Function(x) x.Cantidad)
                                    vDisponibleStockEnPres = lBeStockExistenteTmp.Where(Function(x) x.IdPresentacion <> 0).Sum(Function(x) x.Cantidad)

                                    vCantidadStockZonaNoPicking = lBeStockExistenteZonasNoPicking.Sum(Function(x) x.Cantidad)
                                    vCantidadStockZonaPicking = lBeStockExistenteZonaPicking.Sum(Function(x) x.Cantidad)

                                    vCantidadTotalStock = vCantidadStockZonaNoPicking + vCantidadStockZonaPicking

                                    If (vCantidadSolicitadaPedido > vCantidadStock) AndAlso (vCantidadSolicitadaPedido > vCantidadTotalStock) Then

                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                            pCantidadDisponibleStock = vCantidadStock

                                            Throw New Exception(String.Format("Error_202212140140E:  {0} Código:  {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                               BeProducto.Codigo,
                                                                               vCantidadSolicitadaPedido,
                                                                               vCantidadTotalStock))
                                        Else

                                            If (vCantidadSolicitadaPedido > vDisponibleStockEnPres) Then
                                                If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                                    Throw New Exception(String.Format("ERROR_202212140132D: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                       BeProducto.Codigo,
                                                                       vCantidadSolicitadaPedido,
                                                                       vCantidadTotalStock))
                                                End If
                                            Else
                                                vCantidadSolicitadaPedido = vCantidadStock
                                            End If
                                        End If
                                    End If
                                End If

                            Else

                                If Not BePresentacionDefecto Is Nothing Then
                                    If (BePresentacionDefecto.Factor > 0) Then
                                        vCantidadStockEnPres = Math.Round(vCantidadStock / BePresentacionDefecto.Factor, 6)
                                        vCantidadSolicitadaPedidoEnPres = vCantidadSolicitadaPedido
                                        If vCantidadSolicitadaPedidoEnPres > vCantidadStockEnPres Then
                                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                                Throw New Exception(vbNewLine & String.Format("ERROR_202212140132B: {0} Código: {1} Sol: {2} Disp: {3}", clsDalEx.ErrorS0002A, BeProducto.Codigo, vCantidadSolicitadaPedido, vCantidadStockEnPres))
                                            Else
                                                vCantidadCompletada = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                        Else

                            vCantidadStockZonaNoPicking = lBeStockExistenteZonasNoPicking.Sum(Function(x) x.Cantidad)
                            vCantidadStockZonaPicking = lBeStockExistenteZonaPicking.Sum(Function(x) x.Cantidad)
                            vCantidadTotalStock = vCantidadStockZonaNoPicking + vCantidadStockZonaPicking

                            If (vCantidadSolicitadaPedido > vCantidadStock) AndAlso (vCantidadSolicitadaPedido > vCantidadTotalStock) Then

                                If pBeConfigEnc.Rechazar_pedido_incompleto AndAlso Not pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking Then

                                    pCantidadDisponibleStock = vCantidadDispStock
                                    Throw New Exception(String.Format("Error_202212140140D: {0} Código: {1} Sol: {2} Disp: {3}", clsDalEx.ErrorS0002,
                                                                       BeProducto.Codigo,
                                                                       vCantidadSolicitadaPedido,
                                                                       vCantidadStock))
                                Else

                                    If vCantidadSolicitadaPedido > vDisponibleStockEnPres Then
                                        If pBeConfigEnc.Rechazar_pedido_incompleto Then
                                            Throw New Exception(String.Format("Error_202212140140C: {0} Código: {1} Sol: {2} Disp: {3}",
                                                                              clsDalEx.ErrorS0002A,
                                                                              BeProducto.Codigo,
                                                                              vCantidadSolicitadaPedido,
                                                                              vCantidadStockEnPres))
                                        Else
                                            '#EJC20240808: Asignar en la cantidad solicitada de pedido, lo disponible en el stock recorrido, luego dejar el pendiente para una llamada recursiva u otro ciclo.
                                            '#CKFK20230211 Aqregué esta validación para que no me cambie la cantidad reservada
                                            If pStockResSolicitud.IdPresentacion <> 0 Then
                                                vCantidadSolicitadaPedido = vDisponibleStockEnPres
                                            End If
                                        End If

                                    Else
                                        '#EJC202212140117: Aquí va a nacer otra condición faltante. (caso de unidades).
                                        vCantidadSolicitadaPedido = vDisponibleStockEnPres
                                    End If

                                End If

                            End If

                        End If

                        '#EJC20180614: Cantidad total disponible en stock.
                        vCantidadDispStock = lBeStockExistente.Sum(Function(x) x.Cantidad)

                        '#CKFK20221116 Por lo que he debuggeado aquí se colocan las cantidades tal como se piden en el pedido
                        vCantidadPendiente = vCantidadSolicitadaPedido

                        If IdProducto = 0 Then
                            IdProducto = BeProducto.IdProducto
                        End If

ANALIZAR_FECHAS_DE_VENCIMIENTO:

#Region "Recalcular stocks y fechas de vencimiento"

                        If lBeStockExistente IsNot Nothing Then

                            If lBeStockExistente.Count = 0 Then

                                '#CKFK20231104 Movi esto para dentro de estos if porque si tengo stock no debo recacular
                                If pStockResSolicitud.IdPresentacion <> 0 Then
                                    If Not pStockResBusquedaParaExplosion Is Nothing Then
                                        If pStockResBusquedaParaExplosion.IdProductoBodega = 0 Then
                                            clsPublic.CopyObject(pStockResSolicitud, pStockResBusquedaParaExplosion)
                                            pStockResBusquedaParaExplosion = pStockResSolicitud.Clone()
                                            pStockResBusquedaParaExplosion.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                        End If
                                    End If
                                End If

                                '#EJC202309271639: Se busca explosionar primero de zonas de picking.
                                lBeStockExistenteZonaPicking = clsLnStock.lStock(IIf(vBusquedaEnUmBas,
                                                                                 pStockResSolicitud,
                                                                                 pStockResBusquedaParaExplosion),
                                                                                 BeProducto,
                                                                                 DiasVencimiento,
                                                                                 pBeConfigEnc,
                                                                                 lConnection,
                                                                                 ltransaction,
                                                                                 False,
                                                                                 True,
                                                                                 pTarea_Reabasto,
                                                                                 pEs_Devolucion)

                                Restar_Stock_Reservado(lBeStockExistenteZonaPicking,
                                                       pBeConfigEnc,
                                                       lConnection,
                                                       ltransaction)

                                lBeStockExistenteZonaPicking = lBeStockExistenteZonaPicking.FindAll(Function(x) x.Cantidad > 0)

                                lBeStockExistenteZonasNoPicking = clsLnStock.lStock(IIf(vBusquedaEnUmBas,
                                                                  pStockResSolicitud,
                                                                  pStockResBusquedaParaExplosion),
                                                              BeProducto,
                                                              DiasVencimiento,
                                                              pBeConfigEnc,
                                                              lConnection,
                                                              ltransaction,
                                                              True,
                                                              True,
                                                              pTarea_Reabasto,
                                                              pEs_Devolucion)

                                Restar_Stock_Reservado(lBeStockExistenteZonasNoPicking,
                                                       pBeConfigEnc,
                                                       lConnection,
                                                       ltransaction)

                                lBeStockExistenteZonasNoPicking = lBeStockExistenteZonasNoPicking.FindAll(Function(x) x.Cantidad > 0)

                            End If

                        End If

#End Region

                        Restar_Stock_Reservado(lBeStockExistente,
                                               pBeConfigEnc,
                                               lConnection,
                                               ltransaction)

                        If lBeStockExistente.Count > 0 Then
                            lBeStockExistente = lBeStockExistente.FindAll(Function(x) x.Cantidad > 0)
                        End If

                        vFechaMinimaVenceZonaPicking = vFechaDefecto
                        vFechaMinimaVenceZonaALM = vFechaDefecto
                        vVenceMinimaPickingCompletoClavaud = vFechaDefecto
                        vVenceMinimaPickingInCompletoClavaud = vFechaDefecto

                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                         DiasVencimiento,
                                                                                         pBeConfigEnc,
                                                                                         lConnection,
                                                                                         ltransaction,
                                                                                         BeProducto,
                                                                                         pTarea_Reabasto,
                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                         vFechaMinimaVenceZonaALM,
                                                                                         lBeStockExistente,
                                                                                         BePresentacionDefecto)

                        vFechaMinima = FechaMinimaVenceStock

                        If (vFechaMinimaVenceZonaALM.Date > vFechaMinimaVenceZonaPicking.Date) AndAlso
                            Not (vFechaMinimaVenceZonaPicking.Date = vFechaDefecto) Then

                            If Not lBeStockExistenteZonaPicking Is Nothing Then

                                If lBeStockExistenteZonaPicking.Count > 0 Then

                                    lBeStockExistente = lBeStockExistenteZonaPicking

                                    Restar_Stock_Reservado(lBeStockExistente,
                                                           pBeConfigEnc,
                                                           lConnection,
                                                           ltransaction)

                                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()
                                    vFechaMinima = lBeStockExistente.Min(Function(x) x.Fecha_vence)

                                End If

                            End If

                        Else

                            If (vFechaMinimaVenceZonaALM.Date > vFechaDefecto) AndAlso
                                (vFechaMinimaVenceZonaPicking > vFechaDefecto) AndAlso
                                (vFechaMinimaVenceZonaPicking > vFechaMinimaVenceZonaALM.Date AndAlso
                                vFechaMinimaVenceZonaALM.Date > vFechaDefecto) Then
                                If Not lBeStockExistenteZonasNoPicking Is Nothing Then
                                    If lBeStockExistenteZonasNoPicking.Count > 0 Then
                                        lBeStockExistente = lBeStockExistenteZonasNoPicking
                                    End If
                                End If
                            ElseIf vFechaMinimaVenceZonaPicking > vFechaDefecto Then
                                If Not lBeStockExistenteZonaPicking Is Nothing Then
                                    If lBeStockExistenteZonaPicking.Count > 0 Then
                                        lBeStockExistente = lBeStockExistenteZonaPicking
                                    End If
                                End If
                            End If

                        End If

                        If pBeConfigEnc.Conservar_Zona_Picking_Clavaud Then

                            '#CKFK20240731 Cambié la lista de donde busca Clavaud porque debe ser en las zonas de no picking lBeStockExistenteZonasNoPicking
                            'La lista lBeStockExistente tiene la zona de picking
                            vCantPVConPalletsCompletosClavaud = lBeStockExistenteZonasNoPicking.FindAll(Function(x) x.Pallet_Completo = True _
                                                                                           AndAlso x.UbicacionPicking = False _
                                                                                           AndAlso x.Fecha_vence <= FechaMinimaVenceStock _
                                                                                           AndAlso x.UbicacionNivel > 0).Count()

                            vCantPVConPalletsInCompletosClavaud = lBeStockExistenteZonasNoPicking.FindAll(Function(x) x.Pallet_Completo = False _
                                                                                           AndAlso x.UbicacionPicking = False _
                                                                                           AndAlso x.Fecha_vence <= FechaMinimaVenceStock _
                                                                                           AndAlso x.UbicacionNivel > 0).Count()


                            'Reservo pallets completos que no están en zonas de picking (generalmente rack o piso)
                            lBeStockConPalletsCompletosClavaud = lBeStockExistenteZonasNoPicking.FindAll(Function(x) x.Pallet_Completo = True _
                                                                                           AndAlso x.UbicacionPicking = False _
                                                                                           AndAlso x.Fecha_vence <= FechaMinimaVenceStock _
                                                                                           AndAlso x.UbicacionNivel > 0)

                            'Reservo pallets incompletos (Producidos parcialmente o por proceso) de zonas que no son de picing.
                            lBeStockConPalletsInCompletosClavaud = lBeStockExistenteZonasNoPicking.FindAll(Function(x) x.Pallet_Completo = False _
                                                                                           AndAlso x.UbicacionPicking = False _
                                                                                           AndAlso x.Fecha_vence <= FechaMinimaVenceStock _
                                                                                           AndAlso x.UbicacionNivel > 0)

                            If lBeStockConPalletsCompletosClavaud.Count > 0 Then
                                vVenceMinimaPickingCompletoClavaud = lBeStockConPalletsCompletosClavaud.Min(Function(x) x.Fecha_vence)
                            End If

                            If lBeStockConPalletsInCompletosClavaud.Count > 0 Then
                                vVenceMinimaPickingInCompletoClavaud = lBeStockConPalletsInCompletosClavaud.Min(Function(x) x.Fecha_vence)
                            End If

                            lBeStockZonaPicking = lBeStockExistente.Where(Function(x) x.UbicacionPicking = True AndAlso x.Cantidad > 0).OrderBy(Function(x) x.Fecha_vence).ToList()
                            lBeStockZonasNoPicking = lBeStockExistente.Where(Function(x) x.UbicacionPicking = False AndAlso x.Cantidad > 0).OrderBy(Function(x) x.Fecha_vence).ToList()

                        End If

                        If Iniciar_En = 0 Then
                            If pBeConfigEnc.Conservar_Zona_Picking_Clavaud Then

                                If vVenceMinimaPickingCompletoClavaud > vFechaMinima AndAlso (vVenceMinimaPickingCompletoClavaud.Date > vFechaDefecto) Then
                                    Iniciar_En = 1
                                    'ElseIf vVenceMinimaPickingInCompletoClavaud >= vFechaMinima Then
                                    '    Iniciar_En = 1
                                End If

                                If vFechaMinima >= vVenceMinimaPickingInCompletoClavaud And vVenceMinimaPickingInCompletoClavaud > vFechaDefecto Then
                                    Iniciar_En = 2
                                End If

                            End If
                        End If

                        If vFechaMinimaVenceZonaPicking = vFechaDefecto AndAlso vFechaMinimaVenceZonaALM = vFechaDefecto Then

                            If Not lBeStockExistente Is Nothing Then

                                If lBeStockExistenteZonasNoPicking.Count = 0 Then

                                    pStockResBusquedaParaExplosion = Nothing

                                    If pStockResSolicitud.IdPresentacion = 0 Then '#EJC20230202: Y la primera búsqueda fue en unidades.
                                        If Not BePresentacionDefecto Is Nothing Then '#EJC20230202: Entonces voy a buscar inventario en cajas (con la presentación por defecto si existe)
                                            clsPublic.CopyObject(pStockResSolicitud, pStockResBusquedaParaExplosion)
                                            pStockResBusquedaParaExplosion = pStockResSolicitud.Clone()
                                            pStockResBusquedaParaExplosion.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                            vBusquedaEnUmBas = False
                                        Else
                                            '#CKFK20240320 no aplica lanzar excepcion por esto
                                            ' Throw New Exception("ERROR_202302021127: Se está intentando reservar inventario que requiere explosión pero no está definida la presentación por defecto del producto: " & BeProducto.Codigo)
                                        End If
                                    Else
                                        If Not vEncontroExistenciaEnPresentacion Then

                                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                                Throw New Exception(String.Format("Error_202212140140D: {0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                               BeProducto.Codigo,
                                                                               pStockResSolicitud.Cantidad,
                                                                               0))
                                            Else

                                                If Not vCantidadCompletada AndAlso pStockResSolicitud.IdPresentacion = 0 Then
                                                    '#CKFK20230324 Agregué esta condición para que busque en unidades si la primera búsqueda fue en presentación
                                                    vBusquedaEnUmBas = True
                                                Else
                                                    '#EJC20230724 Se agregó esta condición para que cuando el pedido sea en UMBas y ya no haya existencias en presentación busque en UMBas
                                                    If Not BePedidoDet Is Nothing Then
                                                        If Not vCantidadCompletada AndAlso BePedidoDet.IdPresentacion = 0 Then
                                                            vBusquedaEnUmBas = True
                                                        End If
                                                    End If
                                                End If

                                            End If
                                        Else
                                            vBusquedaEnUmBas = True
                                        End If
                                    End If


                                    '#CKFK20231104 Movi esto para dentro de estos if porque si tengo stock no debo recacular
                                    If pStockResSolicitud.IdPresentacion <> 0 Then
                                        If pStockResBusquedaParaExplosion Is Nothing Then
                                            pStockResBusquedaParaExplosion = New clsBeStock_res()
                                            clsPublic.CopyObject(pStockResSolicitud, pStockResBusquedaParaExplosion)
                                            pStockResBusquedaParaExplosion = pStockResSolicitud.Clone()
                                            pStockResBusquedaParaExplosion.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                        End If
                                    End If

                                    lBeStockExistenteZonasNoPicking = clsLnStock.lStock(IIf(vBusquedaEnUmBas,
                                                                                                      pStockResSolicitud,
                                                                                                      pStockResBusquedaParaExplosion),
                                                                                                  BeProducto,
                                                                                                  DiasVencimiento,
                                                                                                  pBeConfigEnc,
                                                                                                  lConnection,
                                                                                                  ltransaction,
                                                                                                  True,
                                                                                                  True,
                                                                                                  pTarea_Reabasto,
                                                                                                  pEs_Devolucion)

                                    Restar_Stock_Reservado(lBeStockExistenteZonasNoPicking,
                                                           pBeConfigEnc,
                                                           lConnection,
                                                           ltransaction)

                                    If lBeStockExistenteZonasNoPicking IsNot Nothing Then
                                        lBeStockExistenteZonasNoPicking = lBeStockExistenteZonasNoPicking.FindAll(Function(x) x.Cantidad > 0)

                                        If lBeStockExistenteZonasNoPicking.Count > 0 Then
                                            vFechaMinimaVenceZonaALM = lBeStockExistenteZonasNoPicking.Min(Function(x) x.Fecha_vence)
                                            If FechaMinimaVenceStock = vFechaDefecto OrElse FechaMinimaVenceStock > vFechaMinimaVenceZonaALM Then
                                                vFechaMinima = vFechaMinimaVenceZonaALM
                                                Iniciar_En = 4
                                            End If
                                        End If

                                    End If

                                End If

                            End If

                        End If

                        If Iniciar_En = 0 Then
                            If vFechaMinima > vFechaMinimaVenceZonaPicking And vFechaMinimaVenceZonaPicking > vFechaDefecto Then
                                Iniciar_En = 3
                            End If

                            If vFechaMinima > vFechaMinimaVenceZonaALM AndAlso vFechaMinimaVenceZonaALM > vFechaDefecto Then
                                Iniciar_En = 3
                            End If

                            If vFechaMinima > FechaMinimaVenceStock AndAlso FechaMinimaVenceStock > vFechaDefecto Then
                                Iniciar_En = 3
                            End If
                        End If

                        If Iniciar_En = 4 Then
                            If lBeStockExistenteZonasNoPicking.Count > 0 Then
                                lBeStockExistente = lBeStockExistenteZonasNoPicking
                                Iniciar_En = 0
                            End If
                        End If

                        vProcessResult.Add("#MI3_2312201900: Iniciar en: " & Iniciar_En)

                        If BeProducto.Codigo = "WMS223" Then
                            Debug.WriteLine("ESPERA")
                        End If

                        Select Case Iniciar_En
                            Case 1
                                GoTo INICIAR_EN_1 'Tomar pallets completos lBeStockConPalletsCompletosClavaud
                            Case 2
                                GoTo INICIAR_EN_2 'Tomar de pallets incompletos lBeStockConPalletsInCompletosClavaud
                            Case 3
                                GoTo INICIAR_EN_3 'Tomar producto de la zona de picking.
                            Case Else
                                GoTo EJC_202308081248_RESERVAR_DESDE_ULTIMA_LISTA 'Tomar de la lista como la devolvió lStock
                        End Select

INICIAR_EN_1:
                        If Not ListaEstadosDeProceso.Contains(100) Then

                            If pBeConfigEnc.Conservar_Zona_Picking_Clavaud Then

                                For Each vStockOrigen As clsBeStock In lBeStockConPalletsCompletosClavaud

                                    BeStockDestino = New clsBeStock()
                                    clsPublic.CopyObject(vStockOrigen, BeStockDestino)

                                    vCantidadDispStock = Math.Round(vStockOrigen.Cantidad, 6)

                                    If (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                        ListaEstadosDeProceso.Add(100)
                                        GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                    ElseIf (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                        ListaEstadosDeProceso.Add(100)
                                        Exit For
                                    Else
                                        ListaEstadosDeProceso.Add(100)
                                    End If

                                    BeUbicacionEnMemoria = New clsBeBodega_ubicacion With {.IdUbicacion = vStockOrigen.IdUbicacion, .IdBodega = vStockOrigen.IdBodega}

                                    vIndiceUbicacion = lUbicaciones.FindIndex(Function(x) x.IdUbicacion = BeUbicacionEnMemoria.IdUbicacion _
                                                                              AndAlso x.IdBodega = BeUbicacionEnMemoria.IdBodega)

                                    If vIndiceUbicacion <> -1 Then
                                        BeUbicacionStock = New clsBeBodega_ubicacion()
                                        BeUbicacionStock = lUbicaciones(vIndiceUbicacion).Clone()
                                    Else
                                        BeUbicacionStock = New clsBeBodega_ubicacion()
                                        BeUbicacionStock = clsLnBodega_ubicacion.Get_Single_By_IdUbicacion_And_IdBodega(vStockOrigen.IdUbicacion,
                                                                                                                        vStockOrigen.IdBodega,
                                                                                                                        lConnection,
                                                                                                                        ltransaction)
                                        lUbicaciones.Add(BeUbicacionStock.Clone())
                                    End If

                                    If vCantidadDispStock < 0 Then
                                        Throw New Exception("ERROR_202302061300E: La cantidad disponible en stock, reflejó un resultado negativo y no hay fundamento técncio para que eso ocurra (aún), reportar a dsearrollo.")
                                    End If

                                    If vCantidadDispStock > 0 Then
                                        If pStockResSolicitud.IdPresentacion = 0 Then
                                            If pBeConfigEnc.Explosion_Automatica Then
                                                If pBeConfigEnc.Explosion_Automatica_Nivel_Max > 0 Then
                                                    If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then
                                                        If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then
                                                            Continue For
                                                        End If
                                                    End If
                                                Else
                                                    If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then
                                                        Continue For
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then

                                            BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}

                                            vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                            If vIndicePresentacion <> -1 Then
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                            Else
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                If Not BePresentacionDefecto Is Nothing Then
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If
                                            End If

                                            If Not BePresentacionDefecto Is Nothing Then
                                                vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                            End If

                                            Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                            Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                            vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                            vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                            vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                            BeStockRes = New clsBeStock_res
                                            BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                            If BeStockRes.Indicador = "" Then BeStockRes.Indicador = "PED"
                                            BeStockRes.IdBodega = vStockOrigen.IdBodega
                                            BeStockRes.IdStock = vStockOrigen.IdStock
                                            BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                            BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                            If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                            ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                                Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                                If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                                    vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    vCantidadPendiente -= vCantidadPendiente
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                    BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion
                                                Else

                                                    BeStockOriginal = clsLnStock.Get_Single_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                    BeStockDestino.Cantidad = (1 * BePresentacionDefecto.Factor)
                                                    BeStockDestino.Fec_agr = Now
                                                    BeStockDestino.IdPresentacion = 0
                                                    BeStockDestino.Presentacion.IdPresentacion = 0
                                                    BeStockDestino.IdStock = clsLnStock.MaxID(lConnection, ltransaction) + 1
                                                    BeStockRes.IdStock = BeStockDestino.IdStock
                                                    BeStockDestino.No_bulto = 1989
                                                    CantidadStockDestino = BeStockDestino.Cantidad
                                                    'agregar parámetro inferir_decimal_reserva
                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockDestino.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    clsLnStock.Insertar(BeStockDestino,
                                                                        lConnection,
                                                                        ltransaction)

                                                    If vCantidadPendiente > BeStockDestino.Cantidad Then
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente - BeStockDestino.Cantidad
                                                    Else
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    End If

                                                    vStockOrigen.Cantidad = BeStockOriginal.Cantidad - (1 * BePresentacionDefecto.Factor)

                                                    If vStockOrigen.Cantidad > 0 Then
                                                        clsLnStock.Actualizar_Cantidad(vStockOrigen, lConnection, ltransaction)
                                                    Else
                                                        clsLnStock.Eliminar_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                    End If

                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                    BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                    BeStockRes.Lote = vStockOrigen.Lote
                                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                    BeStockRes.Peso = vStockOrigen.Peso
                                                    BeStockRes.Estado = "UNCOMMITED"
                                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                    BeStockRes.Uds_lic_plate = 20220525
                                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.IdPicking = 0
                                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                    BeStockRes.IdDespacho = 0
                                                    BeStockRes.añada = vStockOrigen.Añada
                                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.Host = MaquinaQueSolicita
                                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                    CantidadStockDestino = BeStockRes.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    Insertar(BeStockRes,
                                                             lConnection,
                                                             ltransaction)

                                                    vNombreCasoReservaInternoWMS = "CASO_#1_EJC202310090957"
                                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                    If Not pBeTrasladoDet Is Nothing Then

                                                        If BeStockRes.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            If BePedidoDet.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                            End If
                                                        End If

                                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                     BeProducto,
                                                                                                                     lConnection,
                                                                                                                     ltransaction)
                                                    End If

                                                    lBeStockAReservar.Add(BeStockRes)

                                                    If vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                        'Reservar el remanente en cajas completas.
                                                        vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        BeStockRes = New clsBeStock_res
                                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                        If BeStockRes.Indicador = "" Then
                                                            BeStockRes.Indicador = "PED"
                                                        End If
                                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                        BeStockRes.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                        BeStockRes.Lote = vStockOrigen.Lote
                                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                        BeStockRes.Peso = vStockOrigen.Peso
                                                        BeStockRes.Estado = "UNCOMMITED"
                                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                        BeStockRes.Uds_lic_plate = 20220526 'Marcar el stock reservado para indicar que se tomó a partir de cajas de una solicitud en unidades.
                                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.IdPicking = 0
                                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                        BeStockRes.IdDespacho = 0
                                                        BeStockRes.añada = vStockOrigen.Añada
                                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.Host = MaquinaQueSolicita

                                                        If BeStockRes.Cantidad = 0 Then
                                                            Throw New Exception("Error_202302061305G: La cantidad a reservar no puede ser 0")
                                                        End If

                                                        CantidadStockDestino = BeStockRes.Cantidad

                                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                        Insertar(BeStockRes,
                                                                 lConnection,
                                                                 ltransaction)

                                                        vNombreCasoReservaInternoWMS = "CASO_#2_EJC202310090957"
                                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                        If Not pBeTrasladoDet Is Nothing Then

                                                            If BeStockRes.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                If BePedidoDet.IdPresentacion = 0 Then
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                Else
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                                End If
                                                            End If

                                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                         BeProducto,
                                                                                                                         lConnection,
                                                                                                                         ltransaction)
                                                        End If

                                                        Restar_Stock_Reservado(lBeStockConPalletsCompletosClavaud,
                                                                               pBeConfigEnc,
                                                                               lConnection,
                                                                               ltransaction)

                                                        lBeStockAReservar.Add(BeStockRes)

                                                    End If

                                                    vCantidadCompletada = (vCantidadPendiente = 0)

                                                    vCantidadEnteraTarimasCompletasClavaud -= 1

                                                    If vCantidadCompletada OrElse vCantidadEnteraTarimasCompletasClavaud = 0 Then Exit For

                                                End If

                                            ElseIf vCantidadPendiente > vCantidadDispStock Then

                                                Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)


                                                If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                                    vCantidadAReservarPorIdStock = vCantidadDispStock
                                                    vCantidadPendiente -= vCantidadDispStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                                Else
                                                    Continue For
                                                End If

                                            End If

                                            BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                            BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                            BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                            BeStockRes.Lote = vStockOrigen.Lote
                                            BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                            BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                            BeStockRes.Peso = vStockOrigen.Peso
                                            BeStockRes.Estado = "UNCOMMITED"
                                            BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                            BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                            BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                            BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                            BeStockRes.No_bulto = vStockOrigen.No_bulto
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.IdPicking = 0
                                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                            BeStockRes.IdDespacho = 0
                                            BeStockRes.añada = vStockOrigen.Añada
                                            BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                            BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.Host = MaquinaQueSolicita

                                            If BeStockRes.Cantidad = 0 Then
                                                Throw New Exception("Error_202302061305G: La cantidad a reservar no puede ser 0")
                                            End If

                                            CantidadStockDestino = BeStockRes.Cantidad

                                            vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                            clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                            If Math.Abs(CantidadStockDestino - Fix(CantidadStockDestino)) Then
                                                Throw New Exception("Error_202303101448C: El valor a insertar en stock sería un valor decimal no válido, se prevendrá continuar para evitar inconvenientes en reserva.")
                                            End If

                                            BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                            Insertar(BeStockRes,
                                                     lConnection,
                                                     ltransaction)

                                            vNombreCasoReservaInternoWMS = "CASO_#3_EJC202310090957"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            If Not pBeTrasladoDet Is Nothing Then

                                                If BeStockRes.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    If BePedidoDet.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                    End If
                                                End If

                                                clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                             BeProducto,
                                                                                                             lConnection,
                                                                                                             ltransaction)
                                            End If

                                            vCantidadCompletada = (vCantidadPendiente = 0)
                                            lBeStockAReservar.Add(BeStockRes)

                                            vCantidadEnteraTarimasCompletasClavaud -= 1


                                            Restar_Stock_Reservado(lBeStockConPalletsCompletosClavaud,
                                                                   pBeConfigEnc,
                                                                   lConnection,
                                                                   ltransaction)

                                            If vCantidadCompletada OrElse vCantidadEnteraTarimasCompletasClavaud = 0 Then Exit For

                                        Else

                                            If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                       lConnection,
                                                                                                                                       ltransaction)

                                                If Not BePresentacionDefecto Is Nothing Then

                                                    vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                                    If vIndicePresentacion <> -1 Then
                                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                        BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                    Else
                                                        lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                    End If

                                                    vSolicitudEsEnUMBas = True

                                                End If

                                            ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                                If vIndicePresentacion <> -1 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                Else
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                    clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                    If Not BePresentacionDefecto Is Nothing Then
                                                        lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                    End If
                                                End If

                                                vSolicitudEsEnUMBas = False

                                            End If

                                            If Not BePresentacionDefecto Is Nothing Then
                                                vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                            End If

                                            Split_Decimal(pStockResSolicitud.Cantidad, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                            Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                            If Not BePresentacionDefecto Is Nothing Then

                                                vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                            Else

                                                vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                            End If

                                            If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                                vCantidadPendienteEnPres = vCantidadPendiente
                                                vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                            Else

                                                vCantidadPendienteEnPres = vCantidadPendiente

                                                If pStockResSolicitud.IdPresentacion = 0 Then
                                                    vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                                Else
                                                    vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If

                                                If Not (vStockOrigen.IdPresentacion = 0) Then
                                                    If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                        vConvirtioCantidadSolicitadaEnUmBas = True
                                                    End If
                                                Else
                                                    vCantidadPendiente = vCantidadPendiente
                                                End If

                                            End If

                                            If vSolicitudEsEnUMBas Then

                                                If vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                                    If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso vCantidadDispStockEnPres < BePresentacionDefecto.Factor) Then
                                                        clsLnLog_error_wms.Agregar_Error("#EJC202302081729A: Condición para reserva de unidades alcanzada, impacto desconocido.")
                                                    End If
                                                End If

                                            End If

                                            BeStockRes = New clsBeStock_res
                                            BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                            BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                            BeStockRes.IdBodega = vStockOrigen.IdBodega
                                            BeStockRes.IdStock = vStockOrigen.IdStock
                                            BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                            BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                            If Not vSolicitudEsEnUMBas Then
                                                BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                            End If

                                            If vCantidadPendiente = vCantidadDispStock Then
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            ElseIf vCantidadPendiente < vCantidadDispStock Then
                                                Exit For 'No seguir buscando porque se infiere que no se logrará reservar nada que sea una tarima completa.
                                            ElseIf vCantidadPendiente > vCantidadDispStock Then
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            End If

                                            BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                            BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                            BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                            BeStockRes.Lote = vStockOrigen.Lote
                                            BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                            BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                            BeStockRes.Peso = vStockOrigen.Peso
                                            BeStockRes.Estado = "UNCOMMITED"
                                            BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                            BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                            BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                            BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                            BeStockRes.No_bulto = vStockOrigen.No_bulto
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.IdPicking = 0
                                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                            BeStockRes.IdDespacho = 0
                                            BeStockRes.añada = vStockOrigen.Añada
                                            BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                            BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.Host = MaquinaQueSolicita
                                            BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                            CantidadStockDestino = BeStockRes.Cantidad

                                            vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                            clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                            Insertar(BeStockRes,
                                                     lConnection,
                                                     ltransaction)

                                            vNombreCasoReservaInternoWMS = "CASO_#4_EJC202310090957"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            If Not pBeTrasladoDet Is Nothing Then

                                                If BeStockRes.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    If BePedidoDet.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                    End If
                                                End If

                                                clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                             BeProducto,
                                                                                                             lConnection,
                                                                                                             ltransaction)
                                            End If

                                            vCantidadCompletada = (vCantidadPendiente = 0)
                                            lBeStockAReservar.Add(BeStockRes)

                                            vCantidadEnteraTarimasCompletasClavaud -= 1


                                            Restar_Stock_Reservado(lBeStockConPalletsCompletosClavaud,
                                                                   pBeConfigEnc,
                                                                   lConnection,
                                                                   ltransaction)

                                            If vCantidadCompletada OrElse vCantidadEnteraTarimasCompletasClavaud = 0 Then Exit For

                                        End If

                                    End If

                                Next

INICIAR_EN_2:
                                If Not vCantidadCompletada Then

                                    FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                      DiasVencimiento,
                                                                                                      pBeConfigEnc,
                                                                                                      lConnection,
                                                                                                      ltransaction,
                                                                                                      BeProducto,
                                                                                                      pTarea_Reabasto,
                                                                                                      vFechaMinimaVenceZonaPicking,
                                                                                                      vFechaMinimaVenceZonaALM,
                                                                                                      lBeStockExistente,
                                                                                                      BePresentacionDefecto)

                                    '#EJC202303031732: Condición y parametrizacióin solicitada por Carolina.
                                    If pTarea_Reabasto Then

                                        If pBeConfigEnc.considerar_paletizado_en_reabasto Then

                                            Dim vMensajeNoRellenado As String = "Error_202303031731: La tarea de reabasto para: " & pStockResSolicitud.Codigo_Producto & " no se generará porque no hay tarimas completas y la configuración está activa."

                                            Dim vMsgError As String = String.Format("{0} {1}", MethodBase.GetCurrentMethod.Name(), vMensajeNoRellenado)
                                            clsLnLog_error_wms.Agregar_Error(vMsgError)

                                            XtraMessageBox.Show(vMensajeNoRellenado,
                                                                "Reabasto",
                                                                MessageBoxButtons.OK,
                                                                MessageBoxIcon.Error)

                                            Exit Function

                                        End If

                                    End If

                                    For Each vStockOrigen As clsBeStock In lBeStockConPalletsInCompletosClavaud

                                        BeStockDestino = New clsBeStock()

                                        clsPublic.CopyObject(vStockOrigen, BeStockDestino)

                                        vCantidadDispStock = Math.Round(vStockOrigen.Cantidad, 6)

                                        If (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) AndAlso Not ListaEstadosDeProceso.Contains(101) AndAlso ListaEstadosDeProceso.Contains(100) Then
                                            ListaEstadosDeProceso.Add(101)
                                            GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                        ElseIf (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                            ListaEstadosDeProceso.Add(101)
                                            Exit For
                                        Else
                                            ListaEstadosDeProceso.Add(101)
                                        End If

                                        BeUbicacionEnMemoria = New clsBeBodega_ubicacion With {.IdUbicacion = vStockOrigen.IdUbicacion, .IdBodega = vStockOrigen.IdBodega}

                                        vIndiceUbicacion = lUbicaciones.FindIndex(Function(x) x.IdUbicacion = BeUbicacionEnMemoria.IdUbicacion _
                                                                              AndAlso x.IdBodega = BeUbicacionEnMemoria.IdBodega)

                                        If vIndiceUbicacion <> -1 Then
                                            BeUbicacionStock = New clsBeBodega_ubicacion()
                                            BeUbicacionStock = lUbicaciones(vIndiceUbicacion).Clone()
                                        Else
                                            BeUbicacionStock = New clsBeBodega_ubicacion()
                                            BeUbicacionStock = clsLnBodega_ubicacion.Get_Single_By_IdUbicacion_And_IdBodega(vStockOrigen.IdUbicacion,
                                                                                                                        vStockOrigen.IdBodega,
                                                                                                                        lConnection,
                                                                                                                        ltransaction)
                                            lUbicaciones.Add(BeUbicacionStock.Clone())

                                        End If

                                        If vCantidadDispStock < 0 Then
                                            Throw New Exception("ERROR_202302061300F: La cantidad disponible en stock, reflejó un resultado negativo y no hay fundamento técncio para que eso ocurra (aún), reportar a dsearrollo.")
                                        End If

                                        '#EJC20180620: Si la cantidad de un IdStock es 0, es porque la cantidad reservada es igual a lo disponible, por eso se valida aquí.
                                        If vCantidadDispStock > 0 Then

                                            If pStockResSolicitud.IdPresentacion = 0 Then

                                                If pBeConfigEnc.Explosion_Automatica Then

                                                    If pBeConfigEnc.Explosion_Automatica_Nivel_Max > 0 Then

                                                        If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then

                                                            If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then

                                                                Continue For

                                                            End If

                                                        End If

                                                    Else

                                                        If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then

                                                            Continue For

                                                        End If

                                                    End If

                                                End If

                                            End If

                                            If (pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0) OrElse vBusquedaEnUmBas Then

                                                BePresentacionDefecto = Nothing
                                                vIndicePresentacion = -1

                                                If vStockOrigen.IdPresentacion <> 0 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}
                                                    vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)
                                                End If

                                                If vIndicePresentacion <> -1 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                Else
                                                    If Not vStockOrigen.IdPresentacion = 0 Then
                                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                        BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                        clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                        If Not BePresentacionDefecto Is Nothing Then
                                                            lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                        End If
                                                    Else
                                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                        BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                            lConnection,
                                                                                                                                            ltransaction)

                                                        If Not BePresentacionDefecto Is Nothing Then
                                                            lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                        End If
                                                    End If
                                                End If

                                                If Not BePresentacionDefecto Is Nothing And Not vBusquedaEnUmBas Then
                                                    If Not BePresentacionDefecto.Factor = 0 Then
                                                        vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                                    End If
                                                End If

                                                Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                                Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                                If vCantidadDecimalStockUMBas = 0 AndAlso vBusquedaEnUmBas Then
                                                    vCantidadDecimalStockUMBas = pStockResSolicitud.Cantidad
                                                End If

                                                If Not vBusquedaEnUmBas Then

                                                    If BePresentacionDefecto.Factor > 0 Then
                                                        vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                                    End If

                                                    vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                                    vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                                End If

                                                BeStockRes = New clsBeStock_res
                                                BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                BeStockRes.Indicador = IIf(BeStockRes.Indicador = "", "PED", BeStockRes.Indicador)
                                                BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                BeStockRes.IdStock = vStockOrigen.IdStock
                                                BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                                If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                                    vCantidadAReservarPorIdStock = vCantidadDispStock
                                                    vCantidadPendiente -= vCantidadDispStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                                ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                                    Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                                    If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then

                                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                                        vCantidadPendiente -= vCantidadPendiente
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                        vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                        BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion

                                                    Else

                                                        BeStockOriginal = clsLnStock.Get_Single_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                        BeStockDestino.Cantidad = (1 * BePresentacionDefecto.Factor)
                                                        BeStockDestino.Fec_agr = Now
                                                        BeStockDestino.IdPresentacion = 0
                                                        BeStockDestino.Presentacion.IdPresentacion = 0
                                                        BeStockDestino.IdStock = clsLnStock.MaxID(lConnection, ltransaction) + 1
                                                        BeStockRes.IdStock = BeStockDestino.IdStock
                                                        BeStockDestino.No_bulto = 1989

                                                        CantidadStockDestino = BeStockDestino.Cantidad

                                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                        clsLnStock.Insertar(BeStockDestino,
                                                                            lConnection,
                                                                            ltransaction)



                                                        If vCantidadPendiente > BeStockDestino.Cantidad Then
                                                            vCantidadAReservarPorIdStock = vCantidadPendiente - BeStockDestino.Cantidad
                                                        Else
                                                            vCantidadAReservarPorIdStock = vCantidadPendiente
                                                        End If

                                                        vStockOrigen.Cantidad = BeStockOriginal.Cantidad - (1 * BePresentacionDefecto.Factor)

                                                        If vStockOrigen.Cantidad > 0 Then
                                                            clsLnStock.Actualizar_Cantidad(vStockOrigen, lConnection, ltransaction)
                                                        Else
                                                            clsLnStock.Eliminar_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                        End If

                                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                        vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                        BeStockRes.IdPresentacion = IIf(BeStockRes.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                        BeStockRes.Lote = vStockOrigen.Lote
                                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                        BeStockRes.Peso = vStockOrigen.Peso
                                                        BeStockRes.Estado = "UNCOMMITED"
                                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                        BeStockRes.Uds_lic_plate = 20220525 'Marcar el stock reservado para indicar que se tomó a partir de una caja explosionada.
                                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.IdPicking = 0
                                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                        BeStockRes.IdDespacho = 0
                                                        BeStockRes.añada = vStockOrigen.Añada
                                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.Host = MaquinaQueSolicita
                                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                        CantidadStockDestino = BeStockRes.Cantidad

                                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                        Insertar(BeStockRes,
                                                                 lConnection,
                                                                 ltransaction)

                                                        vNombreCasoReservaInternoWMS = "CASO_#5_EJC202310090957"
                                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                        If Not pBeTrasladoDet Is Nothing Then

                                                            If BeStockRes.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                If BePedidoDet.IdPresentacion = 0 Then
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                Else
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                                End If
                                                            End If

                                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                         BeProducto,
                                                                                                                         lConnection,
                                                                                                                         ltransaction)
                                                        End If

                                                        lBeStockAReservar.Add(BeStockRes)

                                                        Restar_Stock_Reservado(lBeStockConPalletsInCompletosClavaud,
                                                                               pBeConfigEnc,
                                                                               lConnection,
                                                                               ltransaction)

                                                        vCantidadDecimalTarimasCompletasClavaud -= 1

                                                        If vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                            vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                            BeStockRes = New clsBeStock_res
                                                            BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                            BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                                            BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                            BeStockRes.IdStock = vStockOrigen.IdStock
                                                            BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                            BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                                            BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                            BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                            BeStockRes.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                            BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                            BeStockRes.Lote = vStockOrigen.Lote
                                                            BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                            BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                            BeStockRes.Peso = vStockOrigen.Peso
                                                            BeStockRes.Estado = "UNCOMMITED"
                                                            BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                            BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                            BeStockRes.Uds_lic_plate = 20220526 'Marcar el stock reservado para indicar que se tomó a partir de cajas de una solicitud en unidades.
                                                            BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                            BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                            BeStockRes.IdPicking = 0
                                                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                            BeStockRes.IdDespacho = 0
                                                            BeStockRes.añada = vStockOrigen.Añada
                                                            BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                            BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                            BeStockRes.Host = MaquinaQueSolicita
                                                            BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                            If BeStockRes.Cantidad = 0 Then
                                                                Throw New Exception("Error_202302061305A: La cantidad a reservar no puede ser 0")
                                                            End If

                                                            CantidadStockDestino = BeStockRes.Cantidad

                                                            vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                            clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                            Insertar(BeStockRes,
                                                                     lConnection,
                                                                     ltransaction)

                                                            vNombreCasoReservaInternoWMS = "CASO_#6_EJC202310090957"
                                                            vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                            If Not pBeTrasladoDet Is Nothing Then

                                                                If BeStockRes.IdPresentacion = 0 Then
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                Else
                                                                    If BePedidoDet.IdPresentacion = 0 Then
                                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                    Else
                                                                        pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                                    End If
                                                                End If

                                                                clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                             BeProducto,
                                                                                                                             lConnection,
                                                                                                                             ltransaction)
                                                            End If

                                                            lBeStockAReservar.Add(BeStockRes)


                                                            Restar_Stock_Reservado(lBeStockConPalletsInCompletosClavaud,
                                                                               pBeConfigEnc,
                                                                               lConnection,
                                                                               ltransaction)

                                                            vCantidadDecimalTarimasCompletasClavaud -= 1

                                                        End If

                                                        vCantidadCompletada = (vCantidadPendiente = 0)

                                                        If vCantidadCompletada Then Exit For

                                                    End If

                                                ElseIf vCantidadPendiente > vCantidadDispStock Then

                                                    Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                                    If Not vBusquedaEnUmBas Then
                                                        If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                                            vCantidadPendiente -= vCantidadDispStock
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                            vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                                        Else
                                                            Continue For
                                                        End If
                                                    Else
                                                        vCantidadAReservarPorIdStock = vCantidadDispStock
                                                        vCantidadPendiente -= vCantidadDispStock
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                        vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                                    End If

                                                End If

                                                BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                BeStockRes.Lote = vStockOrigen.Lote
                                                BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                BeStockRes.Peso = vStockOrigen.Peso
                                                BeStockRes.Estado = "UNCOMMITED"
                                                BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                                BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                BeStockRes.IdPicking = 0
                                                BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                BeStockRes.IdDespacho = 0
                                                BeStockRes.añada = vStockOrigen.Añada
                                                BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                BeStockRes.Host = MaquinaQueSolicita
                                                BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                If BeStockRes.Cantidad = 0 Then
                                                    Throw New Exception("Error_202302061305A: La cantidad a reservar no puede ser 0")
                                                End If

                                                CantidadStockDestino = BeStockRes.Cantidad

                                                vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                Insertar(BeStockRes,
                                                         lConnection,
                                                         ltransaction)

                                                vNombreCasoReservaInternoWMS = "CASO_#7_EJC202310090957"
                                                vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                If Not pBeTrasladoDet Is Nothing Then

                                                    If BeStockRes.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        If BePedidoDet.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                        End If
                                                    End If

                                                    clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                 BeProducto,
                                                                                                                 lConnection,
                                                                                                                 ltransaction)
                                                End If

                                                vCantidadCompletada = (vCantidadPendiente = 0)

                                                lBeStockAReservar.Add(BeStockRes)


                                                Restar_Stock_Reservado(lBeStockConPalletsInCompletosClavaud,
                                                                       pBeConfigEnc,
                                                                       lConnection,
                                                                       ltransaction)

                                                vCantidadDecimalTarimasCompletasClavaud -= 1

                                                If vCantidadCompletada Then Exit For

                                            Else

                                                'Se pidió en UMBAS y el stock está en UMBAS
                                                If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                       lConnection,
                                                                                                                                       ltransaction)

                                                    If Not BePresentacionDefecto Is Nothing Then

                                                        vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                                        If vIndicePresentacion <> -1 Then
                                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                            BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                        Else
                                                            lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                        End If

                                                        vSolicitudEsEnUMBas = True

                                                    End If


                                                ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                                    If vIndicePresentacion <> -1 Then
                                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                        BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                    Else
                                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                        BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                        clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                        If Not BePresentacionDefecto Is Nothing Then
                                                            lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                        End If
                                                    End If

                                                    vSolicitudEsEnUMBas = False

                                                End If

                                                If Not BePresentacionDefecto Is Nothing Then
                                                    vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                                End If

                                                If Not (pStockResSolicitud.IdPresentacion = 0) Then
                                                    Split_Decimal(pStockResSolicitud.Cantidad, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                                    Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)
                                                Else
                                                    vCantidadDecimalStockUMBas = pStockResSolicitud.Cantidad
                                                End If

                                                If Not BePresentacionDefecto Is Nothing Then
                                                    If Not (pStockResSolicitud.IdPresentacion = 0) Then
                                                        vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                                    End If
                                                Else
                                                    vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                                End If


                                                If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                                    vCantidadPendienteEnPres = vCantidadPendiente
                                                    vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                                Else

                                                    vCantidadPendienteEnPres = vCantidadPendiente

                                                    If pStockResSolicitud.IdPresentacion = 0 Then
                                                        vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                                    Else
                                                        vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                                    End If

                                                    If Not (vStockOrigen.IdPresentacion = 0) Then
                                                        If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                            vConvirtioCantidadSolicitadaEnUmBas = True
                                                        End If
                                                    Else
                                                        vCantidadPendiente = vCantidadPendiente
                                                    End If

                                                End If

                                                If vSolicitudEsEnUMBas Then

                                                    If vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                                        If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso (vCantidadDispStockEnPres < BePresentacionDefecto.Factor)) OrElse (vCantidadPendienteEnPres < vCantidadDispStockEnPres) Then
                                                            Continue For
                                                        End If
                                                    End If

                                                End If

                                                BeStockRes = New clsBeStock_res
                                                BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                                BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                BeStockRes.IdStock = vStockOrigen.IdStock
                                                BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                                If Not vSolicitudEsEnUMBas Then
                                                    BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                End If

                                                If vCantidadPendiente = vCantidadDispStock Then

                                                    vCantidadAReservarPorIdStock = vCantidadDispStock
                                                    vCantidadPendiente -= vCantidadDispStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                ElseIf vCantidadPendiente < vCantidadDispStock Then

                                                    vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                ElseIf vCantidadPendiente > vCantidadDispStock Then

                                                    vCantidadAReservarPorIdStock = vCantidadDispStock
                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                Else
                                                    Continue For
                                                End If

                                                BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                BeStockRes.Lote = vStockOrigen.Lote
                                                BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                BeStockRes.Peso = vStockOrigen.Peso
                                                BeStockRes.Estado = "UNCOMMITED"
                                                BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                                BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                BeStockRes.IdPicking = 0
                                                BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                BeStockRes.IdDespacho = 0
                                                BeStockRes.añada = vStockOrigen.Añada
                                                BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                BeStockRes.Host = MaquinaQueSolicita
                                                BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                If BeStockRes.Cantidad = 0 Then
                                                    Throw New Exception("Error_202302061305B: La cantidad a reservar no puede ser 0")
                                                End If

                                                CantidadStockDestino = BeStockRes.Cantidad

                                                vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                Insertar(BeStockRes,
                                                         lConnection,
                                                         ltransaction)

                                                vNombreCasoReservaInternoWMS = "CASO_#8_EJC202310090957"
                                                vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                If Not pBeTrasladoDet Is Nothing Then

                                                    If BeStockRes.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        If BePedidoDet.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                        End If
                                                    End If

                                                    clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                             BeProducto,
                                                                                                             lConnection,
                                                                                                             ltransaction)
                                                End If

                                                vCantidadCompletada = (vCantidadPendiente = 0)
                                                lBeStockAReservar.Add(BeStockRes)


                                                Restar_Stock_Reservado(lBeStockConPalletsInCompletosClavaud,
                                                                       pBeConfigEnc,
                                                                       lConnection,
                                                                       ltransaction)

                                                If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then
                                                    vCantidadDecimalUMBas = vCantidadPendiente
                                                Else
                                                    If vCantidadPendiente = 0 AndAlso pStockResSolicitud.IdPresentacion = 0 Then
                                                        vCantidadDecimalUMBas = 0
                                                    Else
                                                        If lBeStockConPalletsInCompletosClavaud.Count = 1 AndAlso
                                                        lBeStockConPalletsInCompletosClavaud.Item(0).Cantidad = 0 AndAlso
                                                        vCantidadPendiente > 0 AndAlso vBusquedaEnUmBas Then
                                                            vCantidadDecimalUMBas = vCantidadPendiente
                                                        End If
                                                    End If
                                                End If

                                                vCantidadDecimalTarimasCompletasClavaud -= 1

                                                If vCantidadCompletada Then
                                                    Exit For
                                                End If

                                            End If

                                        End If

                                    Next

                                End If

                            End If

                        End If

                    End If
INICIAR_EN_3:
                    If Not vCantidadCompletada Then

                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                         DiasVencimiento,
                                                                                         pBeConfigEnc,
                                                                                         lConnection,
                                                                                         ltransaction,
                                                                                         BeProducto,
                                                                                         pTarea_Reabasto,
                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                         vFechaMinimaVenceZonaALM,
                                                                                         lBeStockExistente,
                                                                                         BePresentacionDefecto)

                        '#EJC202308081023: Tomar producto de la zona de picking.
                        If lBeStockZonaPicking.Count = 0 Then
                            If lBeStockExistente.Count > 0 Then
                                lBeStockZonaPicking = lBeStockExistente.Where(Function(x) x.UbicacionPicking = True AndAlso x.Cantidad > 0).ToList()
                                '#CKFK20241118 Se agregó el restar stock reservado
                                Restar_Stock_Reservado(lBeStockZonaPicking,
                                                       pBeConfigEnc,
                                                       lConnection,
                                                       ltransaction)
                                If lBeStockZonaPicking.Count > 0 Then
                                    lBeStockZonaPicking = lBeStockZonaPicking.Where(Function(x) x.Cantidad > 0).OrderBy(Function(x) x.Fecha_vence).ToList()
                                End If
                            End If
                        End If

                        '#EJC202308081023: Tomar producto de las zonas de NO picking.
                        lBeStockZonasNoPicking = lBeStockExistente.Where(Function(x) x.UbicacionPicking = False AndAlso x.Cantidad > 0).ToList()
                        If lBeStockZonasNoPicking.Count > 0 Then
                            FechaMinimaVenceStock = lBeStockZonasNoPicking.Min(Function(x) x.Fecha_vence)
                        End If

                    End If


EJC_202308081248_RESERVAR_DESDE_ZONA_PICKING:
#Region "Reservar stock de zona de picking"

                    If Not vCantidadCompletada Then

                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                         DiasVencimiento,
                                                                                         pBeConfigEnc,
                                                                                         lConnection,
                                                                                         ltransaction,
                                                                                         BeProducto,
                                                                                         pTarea_Reabasto,
                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                         vFechaMinimaVenceZonaALM,
                                                                                         lBeStockExistente,
                                                                                         BePresentacionDefecto)

                        If Not ExcepcionFechaVenceEsInferiorEnZonaPicking Then

                            If pStockResSolicitud.IdPresentacion = 0 Then
                                '#EJC: Verificar que en ALM, no existan unidades antes de llevar de PICK.
                                lBeStockExistenteZonasNoPicking = clsLnStock.lStock(pStockResSolicitud,
                                                                                    BeProducto,
                                                                                    DiasVencimiento,
                                                                                    pBeConfigEnc,
                                                                                    lConnection,
                                                                                    ltransaction,
                                                                                    True,
                                                                                    False,
                                                                                    pTarea_Reabasto,
                                                                                    pEs_Devolucion)

                                Restar_Stock_Reservado(lBeStockExistenteZonasNoPicking,
                                                       pBeConfigEnc,
                                                       lConnection,
                                                       ltransaction)

                                lBeStockExistenteZonasNoPicking = lBeStockExistenteZonasNoPicking.Where(Function(x) x.Cantidad > 0).ToList()

                            End If

                            If vSolicitudEsEnUMBas Then
                                If Not lBeStockZonaPicking Is Nothing Then
                                    If lBeStockZonaPicking.Count > 0 Then
                                        lBeStockZonaPicking = lBeStockZonaPicking.OrderBy(Function(x) x.IdPresentacion).ToList()
                                        '#CKFK20241118 Se agregó el restar stock reservado
                                        Restar_Stock_Reservado(lBeStockZonaPicking,
                                                                       pBeConfigEnc,
                                                                       lConnection,
                                                                       ltransaction)
                                        If lBeStockZonaPicking.Count > 0 Then
                                            lBeStockZonaPicking = lBeStockZonaPicking.Where(Function(x) x.Cantidad > 0).OrderBy(Function(x) x.Fecha_vence).ToList()
                                        End If
                                    End If
                                End If
                            ElseIf pStockResSolicitud.IdPresentacion = 0 Then
                                If lBeStockExistenteZonasNoPicking.Count > 0 Then
                                    GoTo EJC_202308081248_RESERVAR_DESDE_ZONA_NO_PICKING1
                                End If
                            End If

                            For Each vStockOrigen As clsBeStock In lBeStockZonaPicking

                                BeStockDestino = New clsBeStock()
                                clsPublic.CopyObject(vStockOrigen, BeStockDestino)

                                vCantidadDispStock = Math.Round(vStockOrigen.Cantidad, 6)

                                If (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) AndAlso Not ListaEstadosDeProceso.Contains(100) AndAlso Not ListaEstadosDeProceso.Contains(101) AndAlso Not ListaEstadosDeProceso.Contains(102) Then
                                    ListaEstadosDeProceso.Add(102)
                                    GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                ElseIf (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                    If Not ListaEstadosDeProceso.Contains(102) Then
                                        ListaEstadosDeProceso.Add(102)
                                    End If
                                    GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                Else
                                    If Not ListaEstadosDeProceso.Contains(102) Then
                                        ListaEstadosDeProceso.Add(102)
                                    End If
                                End If

                                BeUbicacionEnMemoria = New clsBeBodega_ubicacion With {.IdUbicacion = vStockOrigen.IdUbicacion, .IdBodega = vStockOrigen.IdBodega}

                                vIndiceUbicacion = lUbicaciones.FindIndex(Function(x) x.IdUbicacion = BeUbicacionEnMemoria.IdUbicacion _
                                                                              AndAlso x.IdBodega = BeUbicacionEnMemoria.IdBodega)

                                If vIndiceUbicacion <> -1 Then
                                    BeUbicacionStock = New clsBeBodega_ubicacion()
                                    BeUbicacionStock = lUbicaciones(vIndiceUbicacion).Clone()
                                Else
                                    BeUbicacionStock = New clsBeBodega_ubicacion()
                                    BeUbicacionStock = clsLnBodega_ubicacion.Get_Single_By_IdUbicacion_And_IdBodega(vStockOrigen.IdUbicacion,
                                                                                                                    vStockOrigen.IdBodega,
                                                                                                                    lConnection,
                                                                                                                    ltransaction)
                                    If Not BeUbicacionStock Is Nothing Then
                                        lUbicaciones.Add(BeUbicacionStock.Clone())
                                    End If
                                End If

                                If vCantidadDispStock < 0 Then
                                    Throw New Exception("ERROR_202302061300G: La cantidad disponible en stock, reflejó un resultado negativo y no hay fundamento técncio para que eso ocurra (aún), reportar a dsearrollo.")
                                End If

                                If vCantidadDispStock > 0 Then
                                    If pStockResSolicitud.IdPresentacion = 0 Then
                                        If pBeConfigEnc.Explosion_Automatica Then
                                            If pBeConfigEnc.Explosion_Automatica_Nivel_Max > 0 Then
                                                If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then
                                                    If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then
                                                        Continue For
                                                    End If
                                                End If
                                            Else
                                                If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then
                                                    Continue For
                                                End If
                                            End If
                                        End If
                                    End If

                                    If pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then

                                        BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}

                                        vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                        If vIndicePresentacion <> -1 Then
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                        Else
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                            clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                            If Not BePresentacionDefecto Is Nothing Then
                                                lPresentaciones.Add(BePresentacionDefecto.Clone())
                                            End If
                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                        End If

                                        Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                        Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                        vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                        vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                        vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                        BeStockRes = New clsBeStock_res
                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                        If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                        ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                            Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                            If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then

                                                vCantidadAReservarPorIdStock = vCantidadPendiente
                                                vCantidadPendiente -= vCantidadPendiente
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion

                                            Else

                                                BeStockOriginal = clsLnStock.Get_Single_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                BeStockDestino.Cantidad = (1 * BePresentacionDefecto.Factor)
                                                BeStockDestino.Fec_agr = Now
                                                BeStockDestino.IdPresentacion = 0
                                                BeStockDestino.Presentacion.IdPresentacion = 0
                                                BeStockDestino.IdStock = clsLnStock.MaxID(lConnection, ltransaction) + 1
                                                BeStockRes.IdStock = BeStockDestino.IdStock
                                                BeStockDestino.No_bulto = 1989

                                                CantidadStockDestino = BeStockDestino.Cantidad

                                                vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                clsLnStock.Insertar(BeStockDestino,
                                                                        lConnection,
                                                                        ltransaction)

                                                If vCantidadPendiente > BeStockDestino.Cantidad Then
                                                    vCantidadAReservarPorIdStock = vCantidadPendiente - BeStockDestino.Cantidad
                                                Else
                                                    vCantidadAReservarPorIdStock = vCantidadPendiente
                                                End If

                                                '#EJC20220510: Quitar al stock en cajas, las unidades.
                                                vStockOrigen.Cantidad = BeStockOriginal.Cantidad - (1 * BePresentacionDefecto.Factor)

                                                If vStockOrigen.Cantidad > 0 Then
                                                    clsLnStock.Actualizar_Cantidad(vStockOrigen, lConnection, ltransaction)
                                                Else
                                                    clsLnStock.Eliminar_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                End If

                                                vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                BeStockRes.Lote = vStockOrigen.Lote
                                                BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                BeStockRes.Peso = vStockOrigen.Peso
                                                BeStockRes.Estado = "UNCOMMITED"
                                                BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                BeStockRes.Uds_lic_plate = 20220525 'Marcar el stock reservado para indicar que se tomó a partir de una caja explosionada.
                                                BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                BeStockRes.IdPicking = 0
                                                BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                BeStockRes.IdDespacho = 0
                                                BeStockRes.añada = vStockOrigen.Añada
                                                BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                BeStockRes.Host = MaquinaQueSolicita
                                                BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                CantidadStockDestino = BeStockRes.Cantidad

                                                vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                Insertar(BeStockRes,
                                                             lConnection,
                                                             ltransaction)

                                                vNombreCasoReservaInternoWMS = "CASO_#9_EJC202310090957"
                                                vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                If Not pBeTrasladoDet Is Nothing Then

                                                    If BeStockRes.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        If BePedidoDet.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                        End If
                                                    End If

                                                    clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                     BeProducto,
                                                                                                                     lConnection,
                                                                                                                     ltransaction)
                                                End If


                                                Restar_Stock_Reservado(lBeStockZonaPicking,
                                                                       pBeConfigEnc,
                                                                       lConnection,
                                                                       ltransaction)

                                                lBeStockAReservar.Add(BeStockRes)

                                                lBeStockZonaPicking = lBeStockZonaPicking.Where(Function(x) x.Cantidad > 0).ToList()

                                                FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                 DiasVencimiento,
                                                                                                                 pBeConfigEnc,
                                                                                                                 lConnection,
                                                                                                                 ltransaction,
                                                                                                                 BeProducto,
                                                                                                                 pTarea_Reabasto,
                                                                                                                 vFechaMinimaVenceZonaPicking,
                                                                                                                 vFechaMinimaVenceZonaALM,
                                                                                                                 lBeStockExistente,
                                                                                                                 BePresentacionDefecto)

                                                If vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                    'Reservar el remanente en cajas completas.
                                                    vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                    BeStockRes = New clsBeStock_res
                                                    BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                    BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                                    BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                    BeStockRes.IdStock = vStockOrigen.IdStock
                                                    BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                    BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                    BeStockRes.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                    BeStockRes.Lote = vStockOrigen.Lote
                                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                    BeStockRes.Peso = vStockOrigen.Peso
                                                    BeStockRes.Estado = "UNCOMMITED"
                                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                    BeStockRes.Uds_lic_plate = 20220526 'Marcar el stock reservado para indicar que se tomó a partir de cajas de una solicitud en unidades.
                                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.IdPicking = 0
                                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                    BeStockRes.IdDespacho = 0
                                                    BeStockRes.añada = vStockOrigen.Añada
                                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.Host = MaquinaQueSolicita
                                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                    CantidadStockDestino = BeStockRes.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    Insertar(BeStockRes,
                                                             lConnection,
                                                             ltransaction)

                                                    vNombreCasoReservaInternoWMS = "CASO_#10_EJC202310090957"
                                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                    If Not pBeTrasladoDet Is Nothing Then

                                                        If BeStockRes.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            If BePedidoDet.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                            End If
                                                        End If

                                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                     BeProducto,
                                                                                                                     lConnection,
                                                                                                                     ltransaction)
                                                    End If

                                                    Restar_Stock_Reservado(lBeStockZonaPicking,
                                                                           pBeConfigEnc,
                                                                           lConnection,
                                                                           ltransaction)

                                                    lBeStockAReservar.Add(BeStockRes)

                                                    lBeStockZonaPicking = lBeStockZonaPicking.Where(Function(x) x.Cantidad > 0).ToList()

                                                    FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                     DiasVencimiento,
                                                                                                                     pBeConfigEnc,
                                                                                                                     lConnection,
                                                                                                                     ltransaction,
                                                                                                                     BeProducto,
                                                                                                                     pTarea_Reabasto,
                                                                                                                     vFechaMinimaVenceZonaPicking,
                                                                                                                     vFechaMinimaVenceZonaALM,
                                                                                                                     lBeStockExistente,
                                                                                                                     BePresentacionDefecto)

                                                End If

                                                vCantidadCompletada = (vCantidadPendiente = 0)

                                                If vCantidadCompletada Then
                                                    Exit For
                                                End If

                                            End If

                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                            Split_Decimal(vCantidadPendienteEnPres,
                                                          vCantidadEnteraSolicitadaPedidoEnPres,
                                                          vCantidadDecimalSolicitadaPedidoEnPres)

                                            If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                '#EJC202310311502_CASO03
                                                Dim vCantPendientePres As Double = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)

                                                If (vCantPendientePres - vCantidadPendienteEnPres) > 0 Then
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                                Else
                                                    vCantidadPendienteEnPres -= vCantPendientePres
                                                End If

                                            Else
                                                Continue For
                                            End If

                                        End If

                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                        '#EJC202408301232: Mi compilador visual dice que hacía falta esta línea aquí y probablemente en otros casos de arriba.
                                        BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                        BeStockRes.Lote = vStockOrigen.Lote
                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                        BeStockRes.Peso = vStockOrigen.Peso
                                        BeStockRes.Estado = "UNCOMMITED"
                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                        BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.IdPicking = 0
                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                        BeStockRes.IdDespacho = 0
                                        BeStockRes.añada = vStockOrigen.Añada
                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.Host = MaquinaQueSolicita
                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                        CantidadStockDestino = BeStockRes.Cantidad

                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                        Insertar(BeStockRes,
                                                 lConnection,
                                                 ltransaction)

                                        vNombreCasoReservaInternoWMS += "CASO_#11_EJC202310090957"
                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                        If Not pBeTrasladoDet Is Nothing Then

                                            If BeStockRes.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                If BePedidoDet.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If
                                            End If

                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                         BeProducto,
                                                                                                         lConnection,
                                                                                                         ltransaction)
                                        End If

                                        Restar_Stock_Reservado(lBeStockZonaPicking,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                        vCantidadCompletada = (vCantidadPendiente = 0)
                                        lBeStockAReservar.Add(BeStockRes)

                                        lBeStockZonaPicking = lBeStockZonaPicking.Where(Function(x) x.Cantidad > 0).ToList()

                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                         DiasVencimiento,
                                                                                                         pBeConfigEnc,
                                                                                                         lConnection,
                                                                                                         ltransaction,
                                                                                                         BeProducto,
                                                                                                         pTarea_Reabasto,
                                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                                         vFechaMinimaVenceZonaALM,
                                                                                                         lBeStockExistente,
                                                                                                         BePresentacionDefecto)
                                        If vCantidadCompletada Then
                                            Exit For
                                        End If

                                    Else
                                        'Se pidió en UMBAS y el stock está en UMBAS
                                        If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                        lConnection,
                                                                                                                                        ltransaction)

                                            If Not BePresentacionDefecto Is Nothing Then

                                                vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                                If vIndicePresentacion <> -1 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                Else
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If

                                                vSolicitudEsEnUMBas = True

                                            End If


                                        ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                            If vIndicePresentacion <> -1 Then
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                            Else
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                If Not BePresentacionDefecto Is Nothing Then
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If
                                            End If

                                            vSolicitudEsEnUMBas = False

                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                        End If

                                        Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                        Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                        Else
                                            vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                        End If

                                        If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                            vCantidadPendienteEnPres = vCantidadPendiente
                                            vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                        Else

                                            vCantidadPendienteEnPres = vCantidadPendiente

                                            If pStockResSolicitud.IdPresentacion = 0 Then
                                                vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                            Else
                                                vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                            End If

                                            If Not (vStockOrigen.IdPresentacion = 0) Then
                                                If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                    vConvirtioCantidadSolicitadaEnUmBas = True
                                                End If
                                            Else
                                                vCantidadPendiente = vCantidadPendiente
                                            End If

                                        End If

                                        If vSolicitudEsEnUMBas Then
                                            If vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                            ElseIf vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                                If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso vCantidadDispStockEnPres < BePresentacionDefecto.Factor) Then
                                                    clsLnLog_error_wms.Agregar_Error("#EJC202302081729: Condición para reserva de unidades alcanzada, impacto desconocido.")
                                                End If
                                            End If
                                        End If

                                        BeStockRes = New clsBeStock_res
                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                        If Not vSolicitudEsEnUMBas Then
                                            BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                        End If

                                        If vCantidadPendiente = vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        ElseIf vCantidadPendiente < vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadPendiente
                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        End If

                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                        BeStockRes.Lote = vStockOrigen.Lote
                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                        BeStockRes.Peso = vStockOrigen.Peso
                                        BeStockRes.Estado = "UNCOMMITED"
                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                        BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.IdPicking = 0
                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                        BeStockRes.IdDespacho = 0
                                        BeStockRes.añada = vStockOrigen.Añada
                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.Host = MaquinaQueSolicita
                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                        CantidadStockDestino = BeStockRes.Cantidad

                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                        Insertar(BeStockRes,
                                                 lConnection,
                                                 ltransaction)

                                        vNombreCasoReservaInternoWMS = "CASO_#12_EJC202310090957"
                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                        If Not pBeTrasladoDet Is Nothing Then

                                            If BeStockRes.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                If BePedidoDet.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If
                                            End If

                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                         BeProducto,
                                                                                                         lConnection,
                                                                                                         ltransaction)
                                        End If

                                        Restar_Stock_Reservado(lBeStockExistente,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                        vCantidadCompletada = (vCantidadPendiente = 0)
                                        lBeStockAReservar.Add(BeStockRes)

                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                          DiasVencimiento,
                                                                                                          pBeConfigEnc,
                                                                                                          lConnection,
                                                                                                          ltransaction,
                                                                                                          BeProducto,
                                                                                                          pTarea_Reabasto,
                                                                                                          vFechaMinimaVenceZonaPicking,
                                                                                                          vFechaMinimaVenceZonaALM,
                                                                                                          lBeStockExistente,
                                                                                                          BePresentacionDefecto)

                                        If vCantidadCompletada Then
                                            Exit For
                                        End If

                                    End If

                                End If

                            Next

                        Else

#Region "Reserverar stock de zona NO Picking"
EJC_202308081248_RESERVAR_DESDE_ZONA_NO_PICKING1:
                            If Not vCantidadCompletada Then

                                For Each vStockOrigen As clsBeStock In lBeStockZonasNoPicking

                                    BeStockDestino = New clsBeStock()
                                    clsPublic.CopyObject(vStockOrigen, BeStockDestino)

                                    vCantidadDispStock = Math.Round(vStockOrigen.Cantidad, 6)

                                    If (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) _
                                        AndAlso Not ListaEstadosDeProceso.Contains(100) _
                                        AndAlso Not ListaEstadosDeProceso.Contains(101) _
                                        AndAlso Not ListaEstadosDeProceso.Contains(102) _
                                        AndAlso Not ListaEstadosDeProceso.Contains(103) Then
                                        ListaEstadosDeProceso.Add(103)
                                        Exit For
                                    ElseIf (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                        ListaEstadosDeProceso.Add(103)
                                        Exit For
                                    Else
                                        ListaEstadosDeProceso.Add(103)
                                    End If

                                    BeUbicacionEnMemoria = New clsBeBodega_ubicacion With {.IdUbicacion = vStockOrigen.IdUbicacion, .IdBodega = vStockOrigen.IdBodega}

                                    vIndiceUbicacion = lUbicaciones.FindIndex(Function(x) x.IdUbicacion = BeUbicacionEnMemoria.IdUbicacion _
                                                                              AndAlso x.IdBodega = BeUbicacionEnMemoria.IdBodega)

                                    If vIndiceUbicacion <> -1 Then
                                        BeUbicacionStock = New clsBeBodega_ubicacion()
                                        BeUbicacionStock = lUbicaciones(vIndiceUbicacion).Clone()
                                    Else
                                        BeUbicacionStock = New clsBeBodega_ubicacion()
                                        BeUbicacionStock = clsLnBodega_ubicacion.Get_Single_By_IdUbicacion_And_IdBodega(vStockOrigen.IdUbicacion,
                                                                                                                        vStockOrigen.IdBodega,
                                                                                                                        lConnection,
                                                                                                                        ltransaction)
                                        If Not BeUbicacionStock Is Nothing Then
                                            lUbicaciones.Add(BeUbicacionStock.Clone())
                                        End If
                                    End If


                                    If vCantidadDispStock < 0 Then
                                        Throw New Exception("ERROR_202302061300G: La cantidad disponible en stock, reflejó un resultado negativo y no hay fundamento técncio para que eso ocurra (aún), reportar a dsearrollo.")
                                    End If

                                    '#EJC20180620: Si la cantidad de un IdStock es 0, es porque la cantidad reservada es igual a lo disponible, por eso se valida aquí.
                                    If vCantidadDispStock > 0 Then

                                        If pStockResSolicitud.IdPresentacion = 0 Then

                                            If pBeConfigEnc.Explosion_Automatica Then

                                                If pBeConfigEnc.Explosion_Automatica_Nivel_Max > 0 Then

                                                    If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then

                                                        If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then

                                                            Continue For

                                                        End If

                                                    End If

                                                Else

                                                    If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then

                                                        Continue For

                                                    End If

                                                End If

                                            End If


                                        End If

                                        If pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then

                                            BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}

                                            vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                            If vIndicePresentacion <> -1 Then
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                            Else
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                If Not BePresentacionDefecto Is Nothing Then
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If
                                            End If

                                            If Not BePresentacionDefecto Is Nothing Then
                                                vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                            End If

                                            Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                            Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                            vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))

                                            vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                            vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                            BeStockRes = New clsBeStock_res
                                            BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion

                                            BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)

                                            BeStockRes.IdBodega = vStockOrigen.IdBodega
                                            BeStockRes.IdStock = vStockOrigen.IdStock
                                            BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                            BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                            If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                            ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                                Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                                If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then

                                                    vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    vCantidadPendiente -= vCantidadPendiente
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                    BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion

                                                Else


                                                    BeStockOriginal = clsLnStock.Get_Single_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)

                                                    BeStockDestino.Cantidad = (1 * BePresentacionDefecto.Factor)
                                                    BeStockDestino.Fec_agr = Now
                                                    BeStockDestino.IdPresentacion = 0
                                                    BeStockDestino.Presentacion.IdPresentacion = 0
                                                    BeStockDestino.IdStock = clsLnStock.MaxID(lConnection, ltransaction) + 1
                                                    BeStockRes.IdStock = BeStockDestino.IdStock
                                                    BeStockDestino.No_bulto = 1989

                                                    CantidadStockDestino = BeStockDestino.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    clsLnStock.Insertar(BeStockDestino,
                                                                        lConnection,
                                                                        ltransaction)

                                                    If vCantidadPendiente > BeStockDestino.Cantidad Then
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente - BeStockDestino.Cantidad
                                                    Else
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    End If

                                                    vStockOrigen.Cantidad = BeStockOriginal.Cantidad - (1 * BePresentacionDefecto.Factor)

                                                    If vStockOrigen.Cantidad > 0 Then
                                                        clsLnStock.Actualizar_Cantidad(vStockOrigen, lConnection, ltransaction)
                                                    Else
                                                        clsLnStock.Eliminar_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                    End If

                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado

                                                    BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                    BeStockRes.Lote = vStockOrigen.Lote
                                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                    BeStockRes.Peso = vStockOrigen.Peso
                                                    BeStockRes.Estado = "UNCOMMITED"
                                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                    BeStockRes.Uds_lic_plate = 20220525 'Marcar el stock reservado para indicar que se tomó a partir de una caja explosionada.
                                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.IdPicking = 0
                                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                    BeStockRes.IdDespacho = 0
                                                    BeStockRes.añada = vStockOrigen.Añada
                                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.Host = MaquinaQueSolicita
                                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                    CantidadStockDestino = BeStockRes.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    Insertar(BeStockRes,
                                                             lConnection,
                                                             ltransaction)

                                                    vNombreCasoReservaInternoWMS = "CASO_#13_EJC202310090957"
                                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                    If Not pBeTrasladoDet Is Nothing Then

                                                        If BeStockRes.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            If BePedidoDet.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                            End If
                                                        End If

                                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                     BeProducto,
                                                                                                                     lConnection,
                                                                                                                     ltransaction)
                                                    End If

                                                    Restar_Stock_Reservado(lBeStockExistente,
                                                                           pBeConfigEnc,
                                                                           lConnection,
                                                                           ltransaction)

                                                    lBeStockAReservar.Add(BeStockRes)

                                                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                                    FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                     DiasVencimiento,
                                                                                                                     pBeConfigEnc,
                                                                                                                     lConnection,
                                                                                                                     ltransaction,
                                                                                                                     BeProducto,
                                                                                                                     pTarea_Reabasto,
                                                                                                                     vFechaMinimaVenceZonaPicking,
                                                                                                                     vFechaMinimaVenceZonaALM,
                                                                                                                     lBeStockExistente,
                                                                                                                     BePresentacionDefecto)

                                                    If vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                        vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        BeStockRes = New clsBeStock_res
                                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion

                                                        If pStockResSolicitud.Indicador = "" Then
                                                            BeStockRes.Indicador = "PED"
                                                        Else
                                                            BeStockRes.Indicador = pStockResSolicitud.Indicador
                                                        End If

                                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                        BeStockRes.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                        BeStockRes.Lote = vStockOrigen.Lote
                                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                        BeStockRes.Peso = vStockOrigen.Peso
                                                        BeStockRes.Estado = "UNCOMMITED"
                                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                        BeStockRes.Uds_lic_plate = 20220526 'Marcar el stock reservado para indicar que se tomó a partir de cajas de una solicitud en unidades.
                                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.IdPicking = 0
                                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                        BeStockRes.IdDespacho = 0
                                                        BeStockRes.añada = vStockOrigen.Añada
                                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.Host = MaquinaQueSolicita
                                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1
                                                        CantidadStockDestino = BeStockRes.Cantidad

                                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                        Insertar(BeStockRes,
                                                                 lConnection,
                                                                 ltransaction)

                                                        vNombreCasoReservaInternoWMS = "CASO_#14_EJC202310090957"
                                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                        If Not pBeTrasladoDet Is Nothing Then

                                                            If BeStockRes.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                If BePedidoDet.IdPresentacion = 0 Then
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                Else
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                                End If
                                                            End If

                                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                         BeProducto,
                                                                                                                         lConnection,
                                                                                                                         ltransaction)
                                                        End If


                                                        Restar_Stock_Reservado(lBeStockExistente,
                                                                               pBeConfigEnc,
                                                                               lConnection,
                                                                               ltransaction)

                                                        lBeStockAReservar.Add(BeStockRes)

                                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                         DiasVencimiento,
                                                                                                                         pBeConfigEnc,
                                                                                                                         lConnection,
                                                                                                                         ltransaction,
                                                                                                                         BeProducto,
                                                                                                                         pTarea_Reabasto,
                                                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                                                         vFechaMinimaVenceZonaALM,
                                                                                                                         lBeStockExistente,
                                                                                                                         BePresentacionDefecto)

                                                    End If

                                                    vCantidadCompletada = (vCantidadPendiente = 0)

                                                    If vCantidadCompletada Then
                                                        Exit For
                                                    End If

                                                End If

                                            ElseIf vCantidadPendiente > vCantidadDispStock Then

                                                Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                                If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                                    vCantidadAReservarPorIdStock = vCantidadDispStock
                                                    vCantidadPendiente -= vCantidadDispStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                                Else
                                                    Continue For
                                                End If

                                            End If

                                            BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                            BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                            BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                            BeStockRes.Lote = vStockOrigen.Lote
                                            BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                            BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                            BeStockRes.Peso = vStockOrigen.Peso
                                            BeStockRes.Estado = "UNCOMMITED"
                                            BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                            BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                            BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                            BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                            BeStockRes.No_bulto = vStockOrigen.No_bulto
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.IdPicking = 0
                                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                            BeStockRes.IdDespacho = 0
                                            BeStockRes.añada = vStockOrigen.Añada
                                            BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                            BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.Host = MaquinaQueSolicita
                                            BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                            CantidadStockDestino = BeStockRes.Cantidad

                                            vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                            clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                            Insertar(BeStockRes,
                                                     lConnection,
                                                     ltransaction)

                                            vNombreCasoReservaInternoWMS = "CASO_#15_EJC202310090957"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            If Not pBeTrasladoDet Is Nothing Then

                                                If BeStockRes.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    If BePedidoDet.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                    End If
                                                End If

                                                clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                             BeProducto,
                                                                                                             lConnection,
                                                                                                             ltransaction)
                                            End If

                                            Restar_Stock_Reservado(lBeStockExistente,
                                                                   pBeConfigEnc,
                                                                   lConnection,
                                                                   ltransaction)

                                            vCantidadCompletada = (vCantidadPendiente = 0)
                                            lBeStockAReservar.Add(BeStockRes)

                                            lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                            FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                              DiasVencimiento,
                                                                                                              pBeConfigEnc,
                                                                                                              lConnection,
                                                                                                              ltransaction,
                                                                                                              BeProducto,
                                                                                                              pTarea_Reabasto,
                                                                                                              vFechaMinimaVenceZonaPicking,
                                                                                                              vFechaMinimaVenceZonaALM,
                                                                                                              lBeStockExistente,
                                                                                                              BePresentacionDefecto)

                                            If vCantidadCompletada Then
                                                Dim Log_20230301_N As String = String.Format("Log_202303011308N: Cantidad_Comletada_202303011301L: Código: {0} Sol: {1} Reservado: {2}. " & vbNewLine,
                                                                                              BeProducto.Codigo,
                                                                                              vCantidadSolicitadaPedido,
                                                                                              BeStockRes.Cantidad)
                                                clsLnLog_error_wms.Agregar_Error(Log_20230301_N)
                                                Exit For
                                            End If

                                        Else
                                            'Se pidió en UMBAS y el stock está en UMBAS
                                            If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                        lConnection,
                                                                                                                                        ltransaction)

                                                If Not BePresentacionDefecto Is Nothing Then

                                                    vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                                    If vIndicePresentacion <> -1 Then
                                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                        BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                    Else
                                                        lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                    End If

                                                    vSolicitudEsEnUMBas = True

                                                End If


                                            ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                                If vIndicePresentacion <> -1 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                Else
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                    clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                    If Not BePresentacionDefecto Is Nothing Then
                                                        lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                    End If
                                                End If

                                                vSolicitudEsEnUMBas = False

                                            End If

                                            If Not BePresentacionDefecto Is Nothing Then
                                                vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                            End If

                                            Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                            Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                            If Not BePresentacionDefecto Is Nothing Then
                                                vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                            Else
                                                vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                            End If

                                            If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                                vCantidadPendienteEnPres = vCantidadPendiente
                                                vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                            Else

                                                vCantidadPendienteEnPres = vCantidadPendiente

                                                If pStockResSolicitud.IdPresentacion = 0 Then
                                                    vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                                Else
                                                    vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If

                                                If Not (vStockOrigen.IdPresentacion = 0) Then
                                                    If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                        vConvirtioCantidadSolicitadaEnUmBas = True
                                                    End If
                                                Else
                                                    vCantidadPendiente = vCantidadPendiente
                                                End If

                                            End If

                                            If vSolicitudEsEnUMBas Then
                                                If vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                                ElseIf vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                                    If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso vCantidadDispStockEnPres < BePresentacionDefecto.Factor) Then
                                                        clsLnLog_error_wms.Agregar_Error("#EJC202302081729: Condición para reserva de unidades alcanzada, impacto desconocido.")
                                                    End If
                                                End If
                                            End If

                                            BeStockRes = New clsBeStock_res
                                            BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                            BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                            BeStockRes.IdBodega = vStockOrigen.IdBodega
                                            BeStockRes.IdStock = vStockOrigen.IdStock
                                            BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                            BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                            If Not vSolicitudEsEnUMBas Then
                                                BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                            End If

                                            If vCantidadPendiente = vCantidadDispStock Then

                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                            ElseIf vCantidadPendiente < vCantidadDispStock Then

                                                vCantidadAReservarPorIdStock = vCantidadPendiente
                                                vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                            ElseIf vCantidadPendiente > vCantidadDispStock Then

                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                            End If

                                            BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                            BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                            BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                            BeStockRes.Lote = vStockOrigen.Lote
                                            BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                            BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                            BeStockRes.Peso = vStockOrigen.Peso
                                            BeStockRes.Estado = "UNCOMMITED"
                                            BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                            BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                            BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                            BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                            BeStockRes.No_bulto = vStockOrigen.No_bulto
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.IdPicking = 0
                                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                            BeStockRes.IdDespacho = 0
                                            BeStockRes.añada = vStockOrigen.Añada
                                            BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                            BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.Host = MaquinaQueSolicita
                                            BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                            CantidadStockDestino = BeStockRes.Cantidad

                                            vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                            clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                            Insertar(BeStockRes,
                                                     lConnection,
                                                     ltransaction)

                                            vNombreCasoReservaInternoWMS = "CASO_#16_EJC202310090957"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            If Not pBeTrasladoDet Is Nothing Then

                                                If BeStockRes.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    If BePedidoDet.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                    End If
                                                End If

                                                clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                             BeProducto,
                                                                                                             lConnection,
                                                                                                             ltransaction)
                                            End If


                                            Restar_Stock_Reservado(lBeStockExistente,
                                                                   pBeConfigEnc,
                                                                   lConnection,
                                                                   ltransaction)

                                            vCantidadCompletada = (vCantidadPendiente = 0)
                                            lBeStockAReservar.Add(BeStockRes)

                                            lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                            FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                              DiasVencimiento,
                                                                                                              pBeConfigEnc,
                                                                                                              lConnection,
                                                                                                              ltransaction,
                                                                                                              BeProducto,
                                                                                                              pTarea_Reabasto,
                                                                                                              vFechaMinimaVenceZonaPicking,
                                                                                                              vFechaMinimaVenceZonaALM,
                                                                                                              lBeStockExistente,
                                                                                                              BePresentacionDefecto)

                                            If vCantidadCompletada Then
                                                Exit For
                                            End If

                                        End If

                                    End If

                                Next

                            End If
#End Region

                        End If

                    End If

#End Region

EJC_202308081248_RESERVAR_DESDE_ZONA_NO_PICKING:
#Region "Reservar stock de zona NO Picking"

                    If Not vCantidadCompletada Then

                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                          DiasVencimiento,
                                                                                          pBeConfigEnc,
                                                                                          lConnection,
                                                                                          ltransaction,
                                                                                          BeProducto,
                                                                                          pTarea_Reabasto,
                                                                                          vFechaMinimaVenceZonaPicking,
                                                                                          vFechaMinimaVenceZonaALM,
                                                                                          lBeStockExistente,
                                                                                          BePresentacionDefecto)

                        If lBeStockZonasNoPicking.Count = 0 Then
                            If lBeStockExistenteZonasNoPicking.Count > 0 Then
                                lBeStockZonasNoPicking = lBeStockExistenteZonasNoPicking
                                Restar_Stock_Reservado(lBeStockZonasNoPicking,
                                                       pBeConfigEnc,
                                                       lConnection,
                                                       ltransaction)

                                If Not lBeStockZonasNoPicking Is Nothing Then
                                    If lBeStockZonasNoPicking.Count > 0 Then
                                        lBeStockZonasNoPicking = lBeStockZonasNoPicking.Where(Function(x) x.UbicacionPicking = False _
                                                                                              AndAlso x.Cantidad > 0).OrderBy(Function(x) x.Fecha_vence).ToList()
                                    End If
                                End If
                            End If
                        Else
                            '#CKFK20241118 Agregué que reste el stock reservado y que se quiten las cantidad =0
                            Restar_Stock_Reservado(lBeStockZonasNoPicking,
                                                   pBeConfigEnc,
                                                   lConnection,
                                                   ltransaction)
                            If lBeStockZonasNoPicking.Count > 0 Then
                                lBeStockZonasNoPicking = lBeStockZonasNoPicking.Where(Function(x) x.UbicacionPicking = False _
                                                                                      AndAlso x.Cantidad > 0).OrderBy(Function(x) x.Fecha_vence).ToList()
                            End If
                        End If

                        If vSolicitudEsEnUMBas Then
                            If Not lBeStockZonasNoPicking Is Nothing Then
                                If lBeStockZonasNoPicking.Count > 0 Then
                                    'lBeStockZonasNoPicking = lBeStockZonasNoPicking.OrderBy(Function(x) x.IdPresentacion).ToList()
                                    lBeStockZonasNoPicking = lBeStockZonasNoPicking.Where(Function(x) x.IdPresentacion = 0).ToList()
                                End If
                            End If
                        End If

                        For Each vStockOrigen As clsBeStock In lBeStockZonasNoPicking

                            BeStockDestino = New clsBeStock()
                            clsPublic.CopyObject(vStockOrigen, BeStockDestino)

                            vCantidadDispStock = Math.Round(vStockOrigen.Cantidad, 6)

                            If (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(100) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(101) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(102) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(103) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(104) Then
                                If Not ListaEstadosDeProceso.Contains(104) Then
                                    ListaEstadosDeProceso.Add(104)
                                End If
                                GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                            ElseIf (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                If Not ListaEstadosDeProceso.Contains(104) Then
                                    ListaEstadosDeProceso.Add(104)
                                    GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                Else
                                    Exit For
                                End If
                            Else
                                If Not ListaEstadosDeProceso.Contains(104) Then
                                    ListaEstadosDeProceso.Add(104)
                                End If
                            End If

                            BeUbicacionEnMemoria = New clsBeBodega_ubicacion With {.IdUbicacion = vStockOrigen.IdUbicacion, .IdBodega = vStockOrigen.IdBodega}

                            vIndiceUbicacion = lUbicaciones.FindIndex(Function(x) x.IdUbicacion = BeUbicacionEnMemoria.IdUbicacion _
                                                                          AndAlso x.IdBodega = BeUbicacionEnMemoria.IdBodega)

                            If vIndiceUbicacion <> -1 Then
                                BeUbicacionStock = New clsBeBodega_ubicacion()
                                BeUbicacionStock = lUbicaciones(vIndiceUbicacion).Clone()
                            Else
                                BeUbicacionStock = New clsBeBodega_ubicacion()
                                BeUbicacionStock = clsLnBodega_ubicacion.Get_Single_By_IdUbicacion_And_IdBodega(vStockOrigen.IdUbicacion,
                                                                                                                    vStockOrigen.IdBodega,
                                                                                                                    lConnection,
                                                                                                                    ltransaction)
                                If Not BeUbicacionStock Is Nothing Then
                                    lUbicaciones.Add(BeUbicacionStock.Clone())
                                End If
                            End If

                            If vCantidadDispStock < 0 Then
                                Throw New Exception("ERROR_202302061300G: La cantidad disponible en stock, reflejó un resultado negativo y no hay fundamento técncio para que eso ocurra (aún), reportar a dsearrollo.")
                            End If

                            '#EJC20180620: Si la cantidad de un IdStock es 0, es porque la cantidad reservada es igual a lo disponible, por eso se valida aquí.
                            If vCantidadDispStock > 0 Then

                                If pStockResSolicitud.IdPresentacion = 0 Then

                                    If pBeConfigEnc.Explosion_Automatica Then

                                        If pBeConfigEnc.Explosion_Automatica_Nivel_Max > 0 Then

                                            If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then
                                                Continue For
                                            End If

                                        End If

                                    End If

                                End If

                                If pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then

                                    BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}

                                    vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                    If vIndicePresentacion <> -1 Then
                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                        BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                    Else
                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                        BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                        clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                        If Not BePresentacionDefecto Is Nothing Then
                                            lPresentaciones.Add(BePresentacionDefecto.Clone())
                                        End If
                                    End If

                                    If Not BePresentacionDefecto Is Nothing Then
                                        vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                    End If

                                    Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                    Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                    vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                    vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                    vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                    '#EJC202309121310: Revisión escensario explosión de cajas por solicitud de unidades, solo en niveles de picking.
                                    If (vStockOrigen.IdPresentacion <> 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                        If pBeConfigEnc.Explosion_Automatica Then

                                            If Not BeUbicacionStock.Ubicacion_picking Then

                                                If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then

                                                    Dim BeUnidadMedida As New clsBeUnidad_medida
                                                    BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                    If Not BeUnidadMedida Is Nothing Then

                                                        '#CKFK20230918 Agregué esta funcionalidad para obtener el disponible en zona de picking
                                                        pStockResSolicitud.IdPresentacion = 0
                                                        lBeStockDisponible = clsLnStock.lStock(pStockResSolicitud,
                                                                                              BeProducto,
                                                                                              DiasVencimiento,
                                                                                              pBeConfigEnc,
                                                                                              lConnection,
                                                                                              ltransaction,
                                                                                              False,
                                                                                              False,
                                                                                              pTarea_Reabasto,
                                                                                              pEs_Devolucion)

                                                        vStockDispZonaPicking = 0

                                                        If lBeStockDisponible IsNot Nothing Then
                                                            If lBeStockDisponible.Count > 0 Then
                                                                Restar_Stock_Reservado(lBeStockDisponible,
                                                                                       pBeConfigEnc,
                                                                                       lConnection,
                                                                                       ltransaction)

                                                                lBeStockDisponible = lBeStockDisponible.Where(Function(x) x.Cantidad > 0).ToList()
                                                                vStockDispZonaPicking = lBeStockDisponible.Sum(Function(x) x.Cantidad)

                                                            End If
                                                        End If

                                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                                            vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202310312158: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona no picking: " & vStockDispZonaPicking
                                                            Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                                            clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)
                                                        Else
                                                            '#CKFK20240116 Agregué este mensaje para ver lo que pasa
                                                            vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202310312158: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona no picking: " & vStockDispZonaPicking
                                                            clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking, lConnection, ltransaction)
                                                        End If

                                                    End If

                                                End If

                                            Else
                                                Dim BeUnidadMedida As New clsBeUnidad_medida
                                                BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                If Not BeUnidadMedida Is Nothing Then

                                                    '#CKFK20230918 Agregué esta funcionalidad para obtener el disponible en zona de picking
                                                    pStockResSolicitud.IdPresentacion = 0
                                                    lBeStockDisponible = clsLnStock.lStock(pStockResSolicitud,
                                                                                           BeProducto,
                                                                                           DiasVencimiento,
                                                                                           pBeConfigEnc,
                                                                                           lConnection,
                                                                                           ltransaction,
                                                                                           False,
                                                                                           False,
                                                                                           pTarea_Reabasto,
                                                                                           pEs_Devolucion)

                                                    vStockDispZonaPicking = 0

                                                    If lBeStockDisponible IsNot Nothing Then
                                                        If lBeStockDisponible.Count > 0 Then
                                                            Restar_Stock_Reservado(lBeStockDisponible,
                                                                                       pBeConfigEnc,
                                                                                       lConnection,
                                                                                       ltransaction)

                                                            lBeStockDisponible = lBeStockDisponible.Where(Function(x) x.Cantidad > 0).ToList()
                                                            vStockDispZonaPicking = lBeStockDisponible.Sum(Function(x) x.Cantidad)

                                                        End If
                                                    End If

                                                    If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                        vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202309120159F: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp: " & vStockDispZonaPicking
                                                        Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                                        clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)

                                                    Else
                                                        Reserva_Stock_From_MI3 = False
                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If

                                    BeStockRes = New clsBeStock_res
                                    BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                    BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                    BeStockRes.IdBodega = vStockOrigen.IdBodega
                                    BeStockRes.IdStock = vStockOrigen.IdStock
                                    BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                    BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                    If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                        vCantidadAReservarPorIdStock = vCantidadDispStock
                                        vCantidadPendiente -= vCantidadDispStock
                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                        vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                    ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                        If BeUbicacionStock.Ubicacion_picking Then

                                            Split_Decimal(vCantidadPendienteEnPres,
                                                          vCantidadEnteraSolicitadaPedidoEnPres,
                                                          vCantidadDecimalSolicitadaPedidoEnPres)


                                            If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then

                                                vCantidadAReservarPorIdStock = vCantidadPendiente
                                                vCantidadPendiente -= vCantidadPendiente
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion

                                            Else

                                                If BeUbicacionStock.Ubicacion_picking Then

                                                    BeStockOriginal = clsLnStock.Get_Single_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                    BeStockDestino.Cantidad = (1 * BePresentacionDefecto.Factor)
                                                    BeStockDestino.Fec_agr = Now
                                                    BeStockDestino.IdPresentacion = 0
                                                    BeStockDestino.Presentacion.IdPresentacion = 0
                                                    BeStockDestino.IdStock = clsLnStock.MaxID(lConnection, ltransaction) + 1
                                                    BeStockRes.IdStock = BeStockDestino.IdStock
                                                    BeStockDestino.No_bulto = 1989

                                                    CantidadStockDestino = BeStockDestino.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    clsLnStock.Insertar(BeStockDestino,
                                                                        lConnection,
                                                                        ltransaction)

                                                    If vCantidadPendiente > BeStockDestino.Cantidad Then
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente - BeStockDestino.Cantidad
                                                    Else
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    End If

                                                    vStockOrigen.Cantidad = BeStockOriginal.Cantidad - (1 * BePresentacionDefecto.Factor)

                                                    If vStockOrigen.Cantidad > 0 Then
                                                        clsLnStock.Actualizar_Cantidad(vStockOrigen, lConnection, ltransaction)
                                                    Else
                                                        clsLnStock.Eliminar_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                    End If

                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                    vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                    BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                    BeStockRes.Lote = vStockOrigen.Lote
                                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                    BeStockRes.Peso = vStockOrigen.Peso
                                                    BeStockRes.Estado = "UNCOMMITED"
                                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                    BeStockRes.Uds_lic_plate = 20220525 'Marcar el stock reservado para indicar que se tomó a partir de una caja explosionada.
                                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.IdPicking = 0
                                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                    BeStockRes.IdDespacho = 0
                                                    BeStockRes.añada = vStockOrigen.Añada
                                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.Host = MaquinaQueSolicita
                                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                    CantidadStockDestino = BeStockRes.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    Insertar(BeStockRes,
                                                             lConnection,
                                                             ltransaction)

                                                    vNombreCasoReservaInternoWMS = "CASO_#17_EJC202310090957"
                                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                    If Not pBeTrasladoDet Is Nothing Then

                                                        If BeStockRes.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            If BePedidoDet.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                            End If
                                                        End If

                                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                         BeProducto,
                                                                                                                         lConnection,
                                                                                                                         ltransaction)
                                                    End If

                                                    Restar_Stock_Reservado(lBeStockExistente,
                                                                           pBeConfigEnc,
                                                                           lConnection,
                                                                           ltransaction)

                                                    lBeStockAReservar.Add(BeStockRes)

                                                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                                    FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                      DiasVencimiento,
                                                                                                                      pBeConfigEnc,
                                                                                                                      lConnection,
                                                                                                                      ltransaction,
                                                                                                                      BeProducto,
                                                                                                                      pTarea_Reabasto,
                                                                                                                      vFechaMinimaVenceZonaPicking,
                                                                                                                      vFechaMinimaVenceZonaALM,
                                                                                                                      lBeStockExistente,
                                                                                                                      BePresentacionDefecto)

                                                    If vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                        vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        BeStockRes = New clsBeStock_res
                                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                        BeStockRes.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                        BeStockRes.Lote = vStockOrigen.Lote
                                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                        BeStockRes.Peso = vStockOrigen.Peso
                                                        BeStockRes.Estado = "UNCOMMITED"
                                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                        BeStockRes.Uds_lic_plate = 20220526 'Marcar el stock reservado para indicar que se tomó a partir de cajas de una solicitud en unidades.
                                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.IdPicking = 0
                                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                        BeStockRes.IdDespacho = 0
                                                        BeStockRes.añada = vStockOrigen.Añada
                                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.Host = MaquinaQueSolicita
                                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                        CantidadStockDestino = BeStockRes.Cantidad

                                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                        Insertar(BeStockRes,
                                                                 lConnection,
                                                                 ltransaction)

                                                        vNombreCasoReservaInternoWMS = "CASO_#18_EJC202310090957"
                                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                        If Not pBeTrasladoDet Is Nothing Then

                                                            If BeStockRes.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                If BePedidoDet.IdPresentacion = 0 Then
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                Else
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                                End If
                                                            End If

                                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                         BeProducto,
                                                                                                                         lConnection,
                                                                                                                         ltransaction)
                                                        End If


                                                        Restar_Stock_Reservado(lBeStockExistente,
                                                                               pBeConfigEnc,
                                                                               lConnection,
                                                                               ltransaction)

                                                        lBeStockAReservar.Add(BeStockRes)

                                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                          DiasVencimiento,
                                                                                                                          pBeConfigEnc,
                                                                                                                          lConnection,
                                                                                                                          ltransaction,
                                                                                                                          BeProducto,
                                                                                                                          pTarea_Reabasto,
                                                                                                                          vFechaMinimaVenceZonaPicking,
                                                                                                                          vFechaMinimaVenceZonaALM,
                                                                                                                          lBeStockExistente,
                                                                                                                          BePresentacionDefecto)

                                                    End If

                                                    vCantidadCompletada = (vCantidadPendiente = 0)

                                                    If vCantidadCompletada Then
                                                        Exit For
                                                    End If

                                                Else

                                                    If vSolicitudEsEnUMBas Then

                                                        If Not ListaEstadosDeProceso.Contains(108) Then
                                                            ListaEstadosDeProceso.Add(108)
                                                            GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                                        Else
                                                            Exit For
                                                        End If
                                                    Else
                                                        Debug.Write("DICE CAROL QUE AUN HAY UNIDADES.")
                                                        Continue For
                                                    End If

                                                End If

                                            End If

                                        Else

                                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                Dim BeUnidadMedida As New clsBeUnidad_medida
                                                BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                If Not BeUnidadMedida Is Nothing Then
                                                    vMensajeNoExplosionEnZonasNoPicking = "#ERROR_20231101: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp: " & vStockDispZonaPicking
                                                Else
                                                    vMensajeNoExplosionEnZonasNoPicking = "#ERROR_20231101: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " Disp: " & vStockDispZonaPicking
                                                End If

                                                clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking, lConnection, ltransaction)

                                                Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)

                                            Else

                                                Exit For

                                            End If

                                        End If

                                    ElseIf vCantidadPendiente > vCantidadDispStock Then


                                        Split_Decimal(vCantidadPendienteEnPres,
                                                          vCantidadEnteraSolicitadaPedidoEnPres,
                                                          vCantidadDecimalSolicitadaPedidoEnPres)

                                        If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                        Else
                                            Continue For
                                        End If

                                    End If

                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                    BeStockRes.Lote = vStockOrigen.Lote
                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                    BeStockRes.Peso = vStockOrigen.Peso
                                    BeStockRes.Estado = "UNCOMMITED"
                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                    BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                    BeStockRes.IdPicking = 0
                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                    BeStockRes.IdDespacho = 0
                                    BeStockRes.añada = vStockOrigen.Añada
                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                    BeStockRes.Host = MaquinaQueSolicita
                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                    CantidadStockDestino = BeStockRes.Cantidad

                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                    Insertar(BeStockRes,
                                             lConnection,
                                             ltransaction)

                                    vNombreCasoReservaInternoWMS = "CASO_#19_EJC202310090957"
                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                    If Not pBeTrasladoDet Is Nothing Then

                                        If BeStockRes.IdPresentacion = 0 Then
                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                        Else
                                            If BePedidoDet.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                            End If
                                        End If

                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                     BeProducto,
                                                                                                     lConnection,
                                                                                                     ltransaction)
                                    End If

                                    Restar_Stock_Reservado(lBeStockExistente,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                    vCantidadCompletada = (vCantidadPendiente = 0)
                                    lBeStockAReservar.Add(BeStockRes)

                                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                    FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                    DiasVencimiento,
                                                                                                    pBeConfigEnc,
                                                                                                    lConnection,
                                                                                                    ltransaction,
                                                                                                    BeProducto,
                                                                                                    pTarea_Reabasto,
                                                                                                    vFechaMinimaVenceZonaPicking,
                                                                                                    vFechaMinimaVenceZonaALM,
                                                                                                    lBeStockExistente,
                                                                                                    BePresentacionDefecto)

                                    If vCantidadCompletada Then
                                        Dim Log_20230301_N As String = String.Format("Log_202303011308N: Cantidad_Comletada_202303011301L: Código: {0} Sol: {1} Reservado: {2}. " & vbNewLine,
                                                                                          BeProducto.Codigo,
                                                                                          vCantidadSolicitadaPedido,
                                                                                          BeStockRes.Cantidad)
                                        clsLnLog_error_wms.Agregar_Error(Log_20230301_N)
                                        Exit For
                                    End If

                                Else

                                    If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                        BePresentacionDefecto = New clsBeProducto_Presentacion()
                                        BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                    lConnection,
                                                                                                                                    ltransaction)

                                        If Not BePresentacionDefecto Is Nothing Then

                                            vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                            If vIndicePresentacion <> -1 Then
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                            Else
                                                lPresentaciones.Add(BePresentacionDefecto.Clone())
                                            End If

                                            vSolicitudEsEnUMBas = True

                                        End If


                                    ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                        If vIndicePresentacion <> -1 Then
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                        Else
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                            clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                            If Not BePresentacionDefecto Is Nothing Then
                                                lPresentaciones.Add(BePresentacionDefecto.Clone())
                                            End If
                                        End If

                                        vSolicitudEsEnUMBas = False

                                    End If

                                    If Not BePresentacionDefecto Is Nothing Then
                                        vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                    End If

                                    Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                    Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                    If Not BePresentacionDefecto Is Nothing Then
                                        vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                    Else
                                        vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                    End If

                                    If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                        vCantidadPendienteEnPres = vCantidadPendiente
                                        vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                    Else

                                        vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)

                                        If pStockResSolicitud.IdPresentacion = 0 Then
                                            vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                        Else
                                            vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                        End If

                                        If Not (vStockOrigen.IdPresentacion = 0) Then
                                            If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                vConvirtioCantidadSolicitadaEnUmBas = True
                                            End If
                                        Else
                                            vCantidadPendiente = vCantidadPendiente
                                        End If

                                    End If

                                    If vSolicitudEsEnUMBas Then
                                        If vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                            If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso vCantidadDispStockEnPres < BePresentacionDefecto.Factor) Then
                                                clsLnLog_error_wms.Agregar_Error("#EJC202302081729: Condición para reserva de unidades alcanzada, impacto desconocido.")
                                            End If
                                        End If
                                    End If

                                    If (vStockOrigen.IdPresentacion <> 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                        If pBeConfigEnc.Explosion_Automatica Then

                                            If Not BeUbicacionStock.Ubicacion_picking Then

                                                If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then

                                                    Dim BeUnidadMedida As New clsBeUnidad_medida
                                                    BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                    If Not BeUnidadMedida Is Nothing Then

                                                        pStockResSolicitud.IdPresentacion = 0
                                                        lBeStockDisponible = clsLnStock.lStock(pStockResSolicitud,
                                                                                              BeProducto,
                                                                                              DiasVencimiento,
                                                                                              pBeConfigEnc,
                                                                                              lConnection,
                                                                                              ltransaction,
                                                                                              False,
                                                                                              False,
                                                                                              pTarea_Reabasto,
                                                                                              pEs_Devolucion)

                                                        vStockDispZonaPicking = 0

                                                        If lBeStockDisponible IsNot Nothing Then
                                                            If lBeStockDisponible.Count > 0 Then
                                                                Restar_Stock_Reservado(lBeStockDisponible,
                                                                                       pBeConfigEnc,
                                                                                       lConnection,
                                                                                       ltransaction)

                                                                lBeStockDisponible = lBeStockDisponible.Where(Function(x) x.Cantidad > 0).ToList()
                                                                vStockDispZonaPicking = lBeStockDisponible.Sum(Function(x) x.Cantidad)

                                                            End If
                                                        End If

                                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                            vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202309120159C: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona picking: " & vStockDispZonaPicking
                                                            Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                                            clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)

                                                        Else
                                                            Reserva_Stock_From_MI3 = False
                                                        End If

                                                    End If

                                                End If

                                            Else

                                                Dim BeUnidadMedida As New clsBeUnidad_medida
                                                BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                If Not BeUnidadMedida Is Nothing Then

                                                    pStockResSolicitud.IdPresentacion = 0
                                                    lBeStockDisponible = clsLnStock.lStock(pStockResSolicitud,
                                                                                              BeProducto,
                                                                                              DiasVencimiento,
                                                                                              pBeConfigEnc,
                                                                                              lConnection,
                                                                                              ltransaction,
                                                                                              False,
                                                                                              False,
                                                                                              pTarea_Reabasto,
                                                                                              pEs_Devolucion)

                                                    vStockDispZonaPicking = 0

                                                    If lBeStockDisponible IsNot Nothing Then
                                                        If lBeStockDisponible.Count > 0 Then
                                                            Restar_Stock_Reservado(lBeStockDisponible,
                                                                                       pBeConfigEnc,
                                                                                       lConnection,
                                                                                       ltransaction)

                                                            lBeStockDisponible = lBeStockDisponible.Where(Function(x) x.Cantidad > 0).ToList()
                                                            vStockDispZonaPicking = lBeStockDisponible.Sum(Function(x) x.Cantidad)

                                                        End If
                                                    End If

                                                    vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202309120159D: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona picking: " & vStockDispZonaPicking
                                                    Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                                    clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)

                                                End If

                                            End If

                                        End If

                                    End If

                                    BeStockRes = New clsBeStock_res
                                    BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                    BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                    BeStockRes.IdBodega = vStockOrigen.IdBodega
                                    BeStockRes.IdStock = vStockOrigen.IdStock
                                    BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                    BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                    If Not vSolicitudEsEnUMBas Then
                                        BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                    Else
                                        'vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor)
                                    End If

                                    If vCantidadPendiente = vCantidadDispStock Then

                                        vCantidadAReservarPorIdStock = vCantidadDispStock
                                        vCantidadPendiente -= vCantidadDispStock
                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                    ElseIf vCantidadPendiente < vCantidadDispStock Then

                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                    ElseIf vCantidadPendiente > vCantidadDispStock Then

                                        vCantidadAReservarPorIdStock = vCantidadDispStock
                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                    End If

                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                    BeStockRes.Lote = vStockOrigen.Lote
                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                    BeStockRes.Peso = vStockOrigen.Peso
                                    BeStockRes.Estado = "UNCOMMITED"
                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                    BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                    BeStockRes.IdPicking = 0
                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                    BeStockRes.IdDespacho = 0
                                    BeStockRes.añada = vStockOrigen.Añada
                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                    BeStockRes.Host = MaquinaQueSolicita
                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                    CantidadStockDestino = BeStockRes.Cantidad

                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                    Insertar(BeStockRes,
                                             lConnection,
                                             ltransaction)

                                    vNombreCasoReservaInternoWMS = "CASO_#20_EJC202310090957"
                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                    If Not pBeTrasladoDet Is Nothing Then

                                        If BeStockRes.IdPresentacion = 0 Then
                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                        Else
                                            If BePedidoDet.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                            End If
                                        End If

                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                     BeProducto,
                                                                                                     lConnection,
                                                                                                     ltransaction)
                                    End If


                                    Restar_Stock_Reservado(lBeStockZonasNoPicking,
                                                           pBeConfigEnc,
                                                           lConnection,
                                                           ltransaction)

                                    vCantidadCompletada = (vCantidadPendiente = 0)
                                    lBeStockAReservar.Add(BeStockRes)

                                    lBeStockZonasNoPicking = lBeStockZonasNoPicking.Where(Function(x) x.Cantidad > 0).ToList()

                                    '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3
                                    FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                      DiasVencimiento,
                                                                                                      pBeConfigEnc,
                                                                                                      lConnection,
                                                                                                      ltransaction,
                                                                                                      BeProducto,
                                                                                                      pTarea_Reabasto,
                                                                                                      vFechaMinimaVenceZonaPicking,
                                                                                                      vFechaMinimaVenceZonaALM,
                                                                                                      lBeStockExistente,
                                                                                                      BePresentacionDefecto)

                                    If vCantidadCompletada Then
                                        Exit For
                                    End If

                                End If

                            End If

                        Next

                        '#EJC202309120404: Identificar si la reserva se hizo o se está buscando umbas (si quedó cantidad pendiente de reserva)
                        If Not vCantidadCompletada AndAlso pStockResSolicitud.IdPresentacion = 0 Then
                            vBusquedaEnUmBas = True
                        Else
                            If Not vCantidadCompletada Then
                                If Not BePedidoDet Is Nothing Then
                                    If Not vCantidadCompletada AndAlso BePedidoDet.IdPresentacion = 0 Then
                                        vBusquedaEnUmBas = True
                                    End If
                                End If
                            Else
                                pListStockResOUT = lBeStockAReservar
                                Reserva_Stock_From_MI3 = True
                                Exit Function
                            End If
                        End If

                    End If

#End Region


EJC_202308081248_RESERVAR_DESDE_ULTIMA_LISTA:

                    If Not vCantidadCompletada Then

                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                         DiasVencimiento,
                                                                                         pBeConfigEnc,
                                                                                         lConnection,
                                                                                         ltransaction,
                                                                                         BeProducto,
                                                                                         pTarea_Reabasto,
                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                         vFechaMinimaVenceZonaALM,
                                                                                         lBeStockExistente,
                                                                                         BePresentacionDefecto)

                        vProcessResult.Add("#MI3_2312201900: RESERVAR_DESDE_ULTIMA_LISTA: Fecha_Minima_Vence_Zona_Picking: " & vFechaMinimaVenceZonaPicking.Date & " Fecha_Minima_Vence_Zona_ALM: " & vFechaMinimaVenceZonaALM)

                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                        If lBeStockExistente.Count = 0 Then
                            If Not lBeStockExistenteZonaPicking.Count = 0 AndAlso vBusquedaEnUmBas Then
                                lBeStockExistente = lBeStockExistenteZonaPicking
                                lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()
                            End If
                        End If

                        If lBeStockExistente.Count > 0 Then

                            If Not vCantidadCompletada AndAlso pStockResSolicitud.IdPresentacion = 0 Then
                                If vSolicitudEsEnUMBas Then
                                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.IdPresentacion = 0).ToList()
                                    If lBeStockExistente.Count = 0 Then
                                        If Not lBeStockExistenteZonaPicking.Count = 0 Then
                                            lBeStockExistente = lBeStockExistenteZonaPicking.FindAll(Function(x) x.Cantidad > 0)
                                        End If
                                    End If
                                End If
                            Else
                                If Not BePedidoDet Is Nothing Then
                                    If Not vCantidadCompletada AndAlso BePedidoDet.IdPresentacion = 0 Then
                                        vBusquedaEnUmBas = True
                                    End If

                                End If
                            End If

                            For Each vStockOrigen As clsBeStock In lBeStockExistente.FindAll(Function(x) Math.Round(x.Cantidad, 6) > 0)

                                BeStockDestino = New clsBeStock()
                                clsPublic.CopyObject(vStockOrigen, BeStockDestino)

                                vCantidadDispStock = Math.Round(vStockOrigen.Cantidad, 6)

                                If (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(100) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(101) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(102) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(103) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(104) _
                                    AndAlso Not ListaEstadosDeProceso.Contains(105) Then
                                    ListaEstadosDeProceso.Add(105)
                                    '#CKFK20240308 Quité el exit for
                                    'Exit For
                                ElseIf (vStockOrigen.Fecha_vence > FechaMinimaVenceStock) Then
                                    '#EJC20240808: Analizar la fecha corta, lo demás no aplica.
                                    'If Not ListaEstadosDeProceso.Contains(105) Then
                                    '    ListaEstadosDeProceso.Add(105)
                                    '    GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                    'Else
                                    '    '#CKFK20240320 Puse este exit for en comentario porque no aplica ese exit for
                                    '    GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                    'End If
                                    '
                                    '#EJC20241104 Se agregó validacion del proceso 105
                                    If Not ListaEstadosDeProceso.Contains(105) Then
                                        If Not (FechaMinimaVenceStock = New Date(1900, 1, 1) AndAlso pBeConfigEnc.Interface_SAP) Then
                                            GoTo ANALIZAR_FECHAS_DE_VENCIMIENTO
                                        End If
                                    End If
                                Else
                                    If Not ListaEstadosDeProceso.Contains(105) Then
                                        ListaEstadosDeProceso.Add(105)
                                    End If
                                End If

                                BeUbicacionEnMemoria = New clsBeBodega_ubicacion With {.IdUbicacion = vStockOrigen.IdUbicacion, .IdBodega = vStockOrigen.IdBodega}

                                vIndiceUbicacion = lUbicaciones.FindIndex(Function(x) x.IdUbicacion = BeUbicacionEnMemoria.IdUbicacion _
                                                                          AndAlso x.IdBodega = BeUbicacionEnMemoria.IdBodega)

                                If vIndiceUbicacion <> -1 Then
                                    BeUbicacionStock = New clsBeBodega_ubicacion()
                                    BeUbicacionStock = lUbicaciones(vIndiceUbicacion).Clone()
                                Else
                                    BeUbicacionStock = New clsBeBodega_ubicacion()
                                    BeUbicacionStock = clsLnBodega_ubicacion.Get_Single_By_IdUbicacion_And_IdBodega(vStockOrigen.IdUbicacion,
                                                                                                                    vStockOrigen.IdBodega,
                                                                                                                    lConnection,
                                                                                                                    ltransaction)
                                    If Not BeUbicacionStock Is Nothing Then
                                        lUbicaciones.Add(BeUbicacionStock.Clone())
                                    End If
                                End If


                                If vCantidadDispStock < 0 Then
                                    Throw New Exception("ERROR_202302061300G: La cantidad disponible en stock, reflejó un resultado negativo y no hay fundamento técncio para que eso ocurra (aún), reportar a dsearrollo.")
                                End If

                                If vCantidadDispStock > 0 AndAlso Not BeBodega.Permitir_Decimales Then

                                    If pStockResSolicitud.IdPresentacion = 0 Then
                                        If pBeConfigEnc.Explosion_Automatica Then
                                            If pBeConfigEnc.Explosion_Automatica_Nivel_Max > 0 Then
                                                If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then
                                                    If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then
                                                        Continue For
                                                    End If
                                                End If
                                            Else
                                                If Not (pBeConfigEnc.Explosion_Automatica_Desde_Ubicacion_Picking AndAlso BeUbicacionStock.Ubicacion_picking) Then
                                                    If Not vBusquedaEnUmBas Then
                                                        Continue For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    If (vStockOrigen.IdPresentacion <> 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                        If pBeConfigEnc.Explosion_Automatica Then

                                            If Not BeUbicacionStock.Ubicacion_picking Then

                                                If Not pBeConfigEnc.Explosion_Automatica_Nivel_Max >= BeUbicacionStock.Nivel Then

                                                    Dim BeUnidadMedida As New clsBeUnidad_medida
                                                    BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                    If Not BeUnidadMedida Is Nothing Then

                                                        pStockResSolicitud.IdPresentacion = 0
                                                        lBeStockDisponible = clsLnStock.lStock(pStockResSolicitud,
                                                                                              BeProducto,
                                                                                              DiasVencimiento,
                                                                                              pBeConfigEnc,
                                                                                              lConnection,
                                                                                              ltransaction,
                                                                                              False,
                                                                                              False,
                                                                                              pTarea_Reabasto,
                                                                                              pEs_Devolucion)

                                                        vStockDispZonaPicking = 0

                                                        If lBeStockDisponible IsNot Nothing Then
                                                            If lBeStockDisponible.Count > 0 Then
                                                                Restar_Stock_Reservado(lBeStockDisponible,
                                                                                       pBeConfigEnc,
                                                                                       lConnection,
                                                                                       ltransaction)

                                                                lBeStockDisponible = lBeStockDisponible.Where(Function(x) x.Cantidad > 0).ToList()
                                                                vStockDispZonaPicking = lBeStockDisponible.Sum(Function(x) x.Cantidad)

                                                            End If
                                                        End If

                                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                            vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202309120159A: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona picking: " & vStockDispZonaPicking
                                                            Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                                            clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)

                                                        Else
                                                            vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202309120159A: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona picking: " & vStockDispZonaPicking
                                                            vProcessResult.Add(vMensajeNoExplosionEnZonasNoPicking)
                                                            Reserva_Stock_From_MI3 = False : Exit For
                                                        End If

                                                    End If
                                                Else
                                                    vProcessResult.Add("#MI3_240115: La explosión automática está activa, la ubicación encontrada no es de picking y la condición de nivel para la explosión no aplica para la ubicación: " & BeUbicacionStock.IdUbicacion & " Explosion_Automatica_Nivel_Max = " & pBeConfigEnc.Explosion_Automatica_Nivel_Max & " y el nivel de la ubicación es: " & BeUbicacionStock.Nivel)
                                                    Continue For
                                                End If

                                            Else

                                                Dim BeUnidadMedida As New clsBeUnidad_medida
                                                BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                If Not BeUnidadMedida Is Nothing Then

                                                    pStockResSolicitud.IdPresentacion = 0
                                                    lBeStockDisponible = clsLnStock.lStock(pStockResSolicitud,
                                                                                           BeProducto,
                                                                                           DiasVencimiento,
                                                                                           pBeConfigEnc,
                                                                                           lConnection,
                                                                                           ltransaction,
                                                                                           False,
                                                                                           False,
                                                                                           pTarea_Reabasto,
                                                                                           pEs_Devolucion)

                                                    vStockDispZonaPicking = 0

                                                    If lBeStockDisponible IsNot Nothing Then

                                                        If lBeStockDisponible.Count > 0 Then

                                                            Restar_Stock_Reservado(lBeStockDisponible,
                                                                                   pBeConfigEnc,
                                                                                   lConnection,
                                                                                   ltransaction)

                                                            lBeStockDisponible = lBeStockDisponible.Where(Function(x) x.Cantidad > 0).ToList()
                                                            vStockDispZonaPicking = lBeStockDisponible.Sum(Function(x) x.Cantidad)

                                                        End If

                                                    End If

                                                    If Not vStockOrigen.UbicacionPicking Then

                                                        If Not Iniciar_En = 0 Then

                                                            vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202309120159B: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo & " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente & " UM: " & BeUnidadMedida.Nombre & " Disp. zona picking: " & vStockDispZonaPicking
                                                            Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                                            clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)

                                                        End If

                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If

                                    If pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then

                                        BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}

                                        vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                        If vIndicePresentacion <> -1 Then
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                        Else
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                            clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                            If Not BePresentacionDefecto Is Nothing Then
                                                lPresentaciones.Add(BePresentacionDefecto.Clone())
                                            End If
                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                        End If

                                        Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                        Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                        vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                        vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                        vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                        BeStockRes = New clsBeStock_res
                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                        If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                        ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                            If BeUbicacionStock.Ubicacion_picking Then

                                                Split_Decimal(vCantidadPendienteEnPres,
                                                              vCantidadEnteraSolicitadaPedidoEnPres,
                                                              vCantidadDecimalSolicitadaPedidoEnPres)

                                                If vCantidadDecimalSolicitadaPedidoEnPres = 0 OrElse vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                    If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then

                                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                                        vCantidadPendiente -= vCantidadPendiente
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                        vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)

                                                        BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion

                                                    Else

                                                        vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                        vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                        vCantidadPendienteEnPres -= vCantidadEnteraSolicitadaPedidoEnPres

                                                        BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion

                                                    End If


                                                Else

                                                    BeStockOriginal = clsLnStock.Get_Single_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)

                                                    Dim cantidadBase As Decimal

                                                    If vCantidadDecimalSolicitadaPedidoEnPres > 0 Then
                                                        ' Verifica si es un número entero
                                                        If vCantidadDecimalSolicitadaPedidoEnPres = Math.Floor(vCantidadDecimalSolicitadaPedidoEnPres) Then
                                                            cantidadBase = vCantidadDecimalSolicitadaPedidoEnPres
                                                        Else
                                                            cantidadBase = 1
                                                        End If
                                                    Else
                                                        cantidadBase = 1
                                                    End If

                                                    BeStockDestino.Cantidad = cantidadBase * BePresentacionDefecto.Factor
                                                    BeStockDestino.Fec_agr = Now
                                                    BeStockDestino.IdPresentacion = 0
                                                    BeStockDestino.Presentacion.IdPresentacion = 0
                                                    BeStockDestino.IdStock = clsLnStock.MaxID(lConnection, ltransaction) + 1
                                                    BeStockRes.IdStock = BeStockDestino.IdStock
                                                    BeStockDestino.No_bulto = 1989

                                                    CantidadStockDestino = BeStockDestino.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    clsLnStock.Insertar(BeStockDestino,
                                                                        lConnection,
                                                                        ltransaction)

                                                    If vCantidadPendiente > BeStockDestino.Cantidad Then
                                                        vCantidadAReservarPorIdStock = BeStockDestino.Cantidad
                                                    Else
                                                        vCantidadAReservarPorIdStock = vCantidadPendiente
                                                    End If

                                                    vStockOrigen.Cantidad = BeStockOriginal.Cantidad - BeStockDestino.Cantidad

                                                    If vStockOrigen.Cantidad > 0 Then
                                                        clsLnStock.Actualizar_Cantidad(vStockOrigen, lConnection, ltransaction)
                                                    Else
                                                        clsLnStock.Eliminar_By_IdStock(vStockOrigen.IdStock, lConnection, ltransaction)
                                                    End If

                                                    vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                    If vBusquedaEnUmBas Then
                                                        vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                                    Else
                                                        vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                    End If

                                                    BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                    BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                                    BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                                    BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                    BeStockRes.Lote = vStockOrigen.Lote
                                                    BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                    BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                    BeStockRes.Peso = vStockOrigen.Peso
                                                    BeStockRes.Estado = "UNCOMMITED"
                                                    BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                    BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                    BeStockRes.Uds_lic_plate = 20220525 'Marcar el stock reservado para indicar que se tomó a partir de una caja explosionada.
                                                    BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                    BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.IdPicking = 0
                                                    BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                    BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                    BeStockRes.IdDespacho = 0
                                                    BeStockRes.añada = vStockOrigen.Añada
                                                    BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                    BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                    BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                    BeStockRes.Host = MaquinaQueSolicita
                                                    BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                    CantidadStockDestino = BeStockRes.Cantidad

                                                    vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                    clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                    Insertar(BeStockRes,
                                                             lConnection,
                                                             ltransaction)

                                                    vNombreCasoReservaInternoWMS = "CASO_#21_EJC202310090957"
                                                    vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                    If Not pBeTrasladoDet Is Nothing Then

                                                        If BeStockRes.IdPresentacion = 0 Then
                                                            pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                        Else
                                                            If BePedidoDet.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                            End If
                                                        End If

                                                        clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                     BeProducto,
                                                                                                                     lConnection,
                                                                                                                     ltransaction)
                                                    End If


                                                    Restar_Stock_Reservado(lBeStockExistente,
                                                                           pBeConfigEnc,
                                                                           lConnection,
                                                                           ltransaction)

                                                    lBeStockAReservar.Add(BeStockRes)


                                                    lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                                    If lBeStockExistente.Count > 0 Then
                                                        '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3
                                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                         DiasVencimiento,
                                                                                                                         pBeConfigEnc,
                                                                                                                         lConnection,
                                                                                                                         ltransaction,
                                                                                                                         BeProducto,
                                                                                                                         pTarea_Reabasto,
                                                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                                                         vFechaMinimaVenceZonaALM,
                                                                                                                         lBeStockExistente,
                                                                                                                         BePresentacionDefecto)
                                                    End If

                                                    If vCantidadEnteraSolicitadaPedidoEnPres > 0 Then

                                                        'If vCantidadPendiente < BePresentacionDefecto.Factor Then
                                                        '    vCantidadAReservarPorIdStock = vCantidadPendiente
                                                        'Else
                                                        '    'Reservar el remanente en cajas completas.
                                                        '    vCantidadAReservarPorIdStock = Math.Round(vCantidadEnteraSolicitadaPedidoEnPres * BePresentacionDefecto.Factor, 6)
                                                        'End If

                                                        ''vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                        'vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                        'vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        BeStockRes = New clsBeStock_res
                                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado

                                                        If vBusquedaEnUmBas AndAlso (vCantidadAReservarPorIdStock <= BePresentacionDefecto.Factor) Then
                                                            BeStockRes.IdPresentacion = 0
                                                            If (vCantidadAReservarPorIdStock < vStockOrigen.Cantidad) AndAlso (vStockOrigen.Cantidad <= BePresentacionDefecto.Factor) Then
                                                                vStockOrigen.IdPresentacion = 0
                                                                vStockOrigen.User_mod = "RES_MI3"
                                                                clsLnStock.Actualizar_Presentacion(vStockOrigen, lConnection, ltransaction)
                                                            End If
                                                        Else
                                                            BeStockRes.IdPresentacion = BePresentacionDefecto.IdPresentacion
                                                        End If

#Region "Explosión por múltiplo"

                                                        Dim vEsMultiplo As Boolean = True
                                                        Dim cantidadMultiplo As Double = 0

                                                        If Not vSolicitudEsEnUMBas Then
                                                            If Not vOrdernarListaStockSinPresentacionPrimero Then
                                                                If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                                    If Not BePresentacionDefecto Is Nothing Then
                                                                        vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                                    End If
                                                                    vConvirtioCantidadSolicitadaEnUmBas = True
                                                                End If
                                                            End If
                                                        Else
                                                            BeStockRes.IdPresentacion = 0
                                                        End If

                                                        If Not BePresentacionDefecto Is Nothing Then
                                                            If BePresentacionDefecto.Factor <> 0 Then
                                                                vEsMultiplo = (vCantidadPendiente Mod BePresentacionDefecto.Factor = 0)
                                                                cantidadMultiplo = (vCantidadPendiente \ BePresentacionDefecto.Factor) * BePresentacionDefecto.Factor
                                                            End If
                                                        End If

                                                        If Not vEsMultiplo Then
                                                            pStockResSolicitud.IdPresentacion = 0
                                                        End If

                                                        If vStockOrigen.IdStock = 6800 Then
                                                            Debug.Write("espera picking")
                                                        End If

                                                        If vCantidadPendiente = vCantidadDispStock Then

                                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                                            vCantidadPendiente -= vCantidadDispStock
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        ElseIf vCantidadPendiente < vCantidadDispStock Then

                                                            If pStockResSolicitud.IdPresentacion <> 0 Then
                                                                vCantidadAReservarPorIdStock = IIf(vEsMultiplo, vCantidadPendiente, cantidadMultiplo)
                                                            Else
                                                                vCantidadAReservarPorIdStock = vCantidadPendiente
                                                            End If
                                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                                            If pStockResSolicitud.IdPresentacion <> 0 Then
                                                                If Not BePresentacionDefecto Is Nothing Then
                                                                    If BePresentacionDefecto.Factor <> 0 Then
                                                                        vEsMultiplo = (vCantidadDispStock Mod BePresentacionDefecto.Factor = 0)
                                                                        cantidadMultiplo = (vCantidadDispStock \ BePresentacionDefecto.Factor) * BePresentacionDefecto.Factor
                                                                        vCantidadDispStock = cantidadMultiplo
                                                                    End If
                                                                End If
                                                            End If


                                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                                        End If

#End Region

                                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                                        BeStockRes.Lote = vStockOrigen.Lote
                                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                                        BeStockRes.Peso = vStockOrigen.Peso
                                                        BeStockRes.Estado = "UNCOMMITED"
                                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                                        BeStockRes.Uds_lic_plate = 20220526 'Marcar el stock reservado para indicar que se tomó a partir de cajas de una solicitud en unidades.
                                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.IdPicking = 0
                                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                                        BeStockRes.IdDespacho = 0
                                                        BeStockRes.añada = vStockOrigen.Añada
                                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                                        BeStockRes.Host = MaquinaQueSolicita
                                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                                        CantidadStockDestino = BeStockRes.Cantidad

                                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                                        Insertar(BeStockRes,
                                                                 lConnection,
                                                                 ltransaction)

                                                        vNombreCasoReservaInternoWMS = "CASO_#22_EJC202310090957"
                                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                        If Not pBeTrasladoDet Is Nothing Then

                                                            If BeStockRes.IdPresentacion = 0 Then
                                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                            Else
                                                                If BePedidoDet.IdPresentacion = 0 Then
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                                Else
                                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                                End If
                                                            End If

                                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                         BeProducto,
                                                                                                                         lConnection,
                                                                                                                         ltransaction)
                                                        End If


                                                        Restar_Stock_Reservado(lBeStockExistente,
                                                                               pBeConfigEnc,
                                                                               lConnection,
                                                                               ltransaction)

                                                        lBeStockAReservar.Add(BeStockRes)

                                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                                         DiasVencimiento,
                                                                                                                         pBeConfigEnc,
                                                                                                                         lConnection,
                                                                                                                         ltransaction,
                                                                                                                         BeProducto,
                                                                                                                         pTarea_Reabasto,
                                                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                                                         vFechaMinimaVenceZonaALM,
                                                                                                                         lBeStockExistente,
                                                                                                                         BePresentacionDefecto)

                                                    End If

                                                    vCantidadCompletada = (vCantidadPendiente = 0)

                                                    If vCantidadCompletada Then
                                                        Exit For
                                                    End If

                                                End If

                                            Else

                                                Dim BeUnidadMedida As New clsBeUnidad_medida
                                                BeUnidadMedida = clsLnUnidad_medida.GetSingle(pStockResSolicitud.IdUnidadMedida, lConnection, ltransaction)

                                                If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                    vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202310312158: No se puede explosionar producto en zonas de no picking para el producto: " & BeProducto.Codigo &
                                                        " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente &
                                                        " UM: " & BeUnidadMedida.Nombre & " Disp. zona no picking: " & vStockDispZonaPicking

                                                    '#EJC202401291004: Dice Carolina que aquí viene o va dependiendo la perspectiva del observador, que 
                                                    pBeTrasladoDet.Process_Result = vMensajeNoExplosionEnZonasNoPicking
                                                    pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                                    clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                                          lConnection,
                                                                                                          ltransaction)

                                                    clsLnLog_error_wms.Agregar_Error(vMensajeNoExplosionEnZonasNoPicking)

                                                    Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)

                                                Else

                                                    vMensajeNoExplosionEnZonasNoPicking = "#ERROR_202401291007: No se puede explosionar producto en zonas de ALM para el producto: " & BeProducto.Codigo &
                                                        " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente &
                                                        " UM: " & BeUnidadMedida.Nombre & " Disp. zona no picking: " & vStockDispZonaPicking

                                                    '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                                    pBeTrasladoDet.Process_Result = vMensajeNoExplosionEnZonasNoPicking
                                                    pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                                    clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                                          lConnection,
                                                                                                          ltransaction)

                                                    '#CKFK20240115 Erik dice que es continue for
                                                    Continue For
                                                End If

                                            End If

                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                            Split_Decimal(vCantidadPendienteEnPres,
                                                          vCantidadEnteraSolicitadaPedidoEnPres,
                                                          vCantidadDecimalSolicitadaPedidoEnPres)

                                            If vCantidadDecimalSolicitadaPedidoEnPres = 0 Then
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                            Else
                                                '#EJC20231031: CASO10
                                                vCantidadAReservarPorIdStock = vCantidadDispStock
                                                vCantidadPendiente -= vCantidadDispStock
                                                vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                            End If

                                        End If

                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                        BeStockRes.Lote = vStockOrigen.Lote
                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                        BeStockRes.Peso = vStockOrigen.Peso
                                        BeStockRes.Estado = "UNCOMMITED"
                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                        BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.IdPicking = 0
                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                        BeStockRes.IdDespacho = 0
                                        BeStockRes.añada = vStockOrigen.Añada
                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.Host = MaquinaQueSolicita
                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                        '#EJC20231031: Viene de un proceso recursivo (probablemente) tratando de reservar en presentación.
                                        If Not vSolicitudEsEnUMBas Then
                                            BeStockRes.IdPresentacion = vStockOrigen.IdPresentacion
                                        End If

                                        CantidadStockDestino = BeStockRes.Cantidad

                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                        Insertar(BeStockRes,
                                                 lConnection,
                                                 ltransaction)

                                        vNombreCasoReservaInternoWMS = "CASO_#23_EJC202310090957"
                                        vMensajeReserva = vNombreCasoReservaInternoWMS + " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                        If Not pBeTrasladoDet Is Nothing Then


                                            If BeStockRes.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                If BePedidoDet.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If
                                            End If

                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                         BeProducto,
                                                                                                         lConnection,
                                                                                                         ltransaction)
                                        End If


                                        Restar_Stock_Reservado(lBeStockExistente,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                        vCantidadCompletada = (vCantidadPendiente = 0)
                                        lBeStockAReservar.Add(BeStockRes)

                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                        If lBeStockExistente.Count > 0 Then
                                            '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3                                            
                                            FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                             DiasVencimiento,
                                                                                                             pBeConfigEnc,
                                                                                                             lConnection,
                                                                                                             ltransaction,
                                                                                                             BeProducto,
                                                                                                             pTarea_Reabasto,
                                                                                                             vFechaMinimaVenceZonaPicking,
                                                                                                             vFechaMinimaVenceZonaALM,
                                                                                                             lBeStockExistente,
                                                                                                             BePresentacionDefecto)
                                        End If

                                        If vCantidadCompletada Then
                                            Exit For
                                        End If

                                    Else

                                        If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                    lConnection,
                                                                                                                                    ltransaction)

                                            If Not BePresentacionDefecto Is Nothing Then

                                                vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                                If vIndicePresentacion <> -1 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                Else
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If

                                                vSolicitudEsEnUMBas = True

                                            End If


                                        ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                            If vIndicePresentacion <> -1 Then
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                            Else
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                If Not BePresentacionDefecto Is Nothing Then
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If
                                            End If

                                            vSolicitudEsEnUMBas = False

                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                        End If

                                        Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                        Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                        Else
                                            vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                        End If

                                        If vCantidadDecimalStockUMBas <> 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then
                                            Dim vMensajePresentacion As String = ""

                                            vMensajePresentacion = "#ERROR_20250615: La cantidad existente no es válida para el producto : " & BeProducto.Codigo &
                                                " Linea: " & No_Linea & " Cantidad: " & vCantidadPendiente &
                                                " Presentación: " & BePresentacionDefecto.Nombre & " Disp. en IdStock: " & vCantidadDecimalStockUMBas &
                                                " Factor: " & BePresentacionDefecto.Factor

                                            '#CKFK20240115 Erik dice que es continue for
                                            Continue For
                                        End If

                                        '#EJC20231019: CASO 15 Se agregó NOT vOrdernarListaStockSinPresentacionPrimero
                                        If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso Not vOrdernarListaStockSinPresentacionPrimero Then

                                            vCantidadPendienteEnPres = vCantidadPendiente
                                            vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                        Else

                                            If vSolicitudEsEnUMBas Then
                                                vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                            Else
                                                vCantidadPendienteEnPres = vCantidadPendiente
                                            End If

                                            If pStockResSolicitud.IdPresentacion = 0 Then
                                                vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                            Else
                                                vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                            End If

                                            If Not (vStockOrigen.IdPresentacion = 0) Then
                                                If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                    If Not vOrdernarListaStockSinPresentacionPrimero Then
                                                        If vSolicitudEsEnUMBas Then
                                                            '#EJC20231019: CASO 15 En vCantidadPendienteEnPres el valor asignado corresponde a las unidades.
                                                            vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                            vConvirtioCantidadSolicitadaEnUmBas = True
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                vCantidadPendiente = vCantidadPendiente
                                            End If

                                        End If

                                        If (vSolicitudEsEnUMBas) Then

                                            If vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                                If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso vCantidadDispStockEnPres < BePresentacionDefecto.Factor) Then
                                                    clsLnLog_error_wms.Agregar_Error("#EJC202302081729: Condición para reserva de unidades alcanzada, impacto desconocido.")
                                                End If
                                            End If
                                        Else
                                            If pStockResSolicitud.IdPresentacion <> 0 AndAlso Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                vConvirtioCantidadSolicitadaEnUmBas = True
                                            End If
                                        End If

                                        BeStockRes = New clsBeStock_res
                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega
                                        BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)

                                        '#EJC20250612: Mejora en el manejo de múltiplos en la reserva.
                                        Dim vEsMultiplo As Boolean = True
                                        Dim cantidadMultiplo As Double = 0

                                        If Not vSolicitudEsEnUMBas Then
                                            If Not vOrdernarListaStockSinPresentacionPrimero Then
                                                If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                    If Not BePresentacionDefecto Is Nothing Then
                                                        vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                    End If
                                                    vConvirtioCantidadSolicitadaEnUmBas = True
                                                End If
                                            Else
                                                BeStockRes.IdPresentacion = 0
                                            End If
                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            If BePresentacionDefecto.Factor <> 0 Then
                                                vEsMultiplo = (vCantidadPendiente Mod BePresentacionDefecto.Factor = 0)
                                                cantidadMultiplo = (vCantidadPendiente \ BePresentacionDefecto.Factor) * BePresentacionDefecto.Factor
                                            End If
                                        End If

                                        If Not vEsMultiplo Then
                                            pStockResSolicitud.IdPresentacion = 0
                                        End If

                                        If vStockOrigen.IdStock = 6800 Then
                                            Debug.Write("espera picking")
                                        End If

                                        If vCantidadPendiente = vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        ElseIf vCantidadPendiente < vCantidadDispStock Then

                                            If pStockResSolicitud.IdPresentacion <> 0 Then
                                                vCantidadAReservarPorIdStock = IIf(vEsMultiplo, vCantidadPendiente, cantidadMultiplo)
                                            Else
                                                vCantidadAReservarPorIdStock = vCantidadPendiente
                                            End If
                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                            If pStockResSolicitud.IdPresentacion <> 0 Then
                                                If Not BePresentacionDefecto Is Nothing Then
                                                    If BePresentacionDefecto.Factor <> 0 Then
                                                        vEsMultiplo = (vCantidadDispStock Mod BePresentacionDefecto.Factor = 0)
                                                        cantidadMultiplo = (vCantidadDispStock \ BePresentacionDefecto.Factor) * BePresentacionDefecto.Factor
                                                        vCantidadDispStock = cantidadMultiplo
                                                    End If
                                                End If
                                            End If


                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        End If

                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                        BeStockRes.Lote = vStockOrigen.Lote
                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                        BeStockRes.Peso = vStockOrigen.Peso
                                        BeStockRes.Estado = "UNCOMMITED"
                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                        BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.IdPicking = 0
                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                        BeStockRes.IdDespacho = 0
                                        BeStockRes.añada = vStockOrigen.Añada
                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.Host = MaquinaQueSolicita
                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1
                                        CantidadStockDestino = BeStockRes.Cantidad

                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                        Insertar(BeStockRes,
                                                 lConnection,
                                                 ltransaction)

                                        vNombreCasoReservaInternoWMS = "CASO_#24_EJC202310090957"
                                        vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                        " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                        " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                        " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                        " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                        " Lote: " & BeStockRes.Lote &
                                                                        " Ubicación: " & BeStockRes.IdUbicacion &
                                                                        " Cantidad: " & BeStockRes.Cantidad &
                                                                        " UmBas: " & BeStockRes.IdUnidadMedida &
                                                                        " Presentacion: " & BeStockRes.Atributo_Variante_1

                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                        If Not pBeTrasladoDet Is Nothing Then

                                            If BeStockRes.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                If BePedidoDet.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If
                                            End If

                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                     BeProducto,
                                                                                                     lConnection,
                                                                                                     ltransaction)
                                        End If

                                        Restar_Stock_Reservado(lBeStockExistente,
                                                           pBeConfigEnc,
                                                           lConnection,
                                                           ltransaction)

                                        vCantidadCompletada = (vCantidadPendiente = 0)
                                        lBeStockAReservar.Add(BeStockRes)


                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                        '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3
                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                     DiasVencimiento,
                                                                                                     pBeConfigEnc,
                                                                                                     lConnection,
                                                                                                     ltransaction,
                                                                                                     BeProducto,
                                                                                                     pTarea_Reabasto,
                                                                                                     vFechaMinimaVenceZonaPicking,
                                                                                                     vFechaMinimaVenceZonaALM,
                                                                                                     lBeStockExistente,
                                                                                                     BePresentacionDefecto)
                                        If vCantidadCompletada Then
                                            Exit For
                                        End If

                                    End If

                                Else

                                    If pStockResSolicitud.IdPresentacion = 0 AndAlso vStockOrigen.IdPresentacion <> 0 Then

                                        BePresentacionDefecto = New clsBeProducto_Presentacion With {.IdPresentacion = vStockOrigen.IdPresentacion}

                                        vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                        If vIndicePresentacion <> -1 Then
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                        Else
                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                            clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                            If Not BePresentacionDefecto Is Nothing Then
                                                lPresentaciones.Add(BePresentacionDefecto.Clone())
                                            End If
                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                        End If

                                        Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                        Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                        vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                        vCantidadPendienteEnPres = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                        vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)

                                        BeStockRes = New clsBeStock_res
                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                        If vCantidadPendienteEnPres = vCantidadDispStockEnPres Then
                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)
                                        ElseIf vCantidadPendienteEnPres < vCantidadDispStockEnPres Then

                                            vCantidadAReservarPorIdStock = vCantidadPendiente
                                            vCantidadPendiente -= vCantidadPendiente
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                            If vCantidadPendiente = 0 Then
                                                vCantidadPendienteEnPres = 0
                                            Else
                                                vCantidadPendienteEnPres -= Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                            End If

                                            BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion
                                            BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                            BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                            If pStockResSolicitud.IdPresentacion = 0 Then
                                                BeStockRes.IdPresentacion = 0
                                            Else
                                                BeStockRes.IdPresentacion = vStockOrigen.Presentacion.IdPresentacion
                                            End If
                                            BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                            BeStockRes.Lote = vStockOrigen.Lote
                                            BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                            BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                            BeStockRes.Peso = vStockOrigen.Peso
                                            BeStockRes.Estado = "UNCOMMITED"
                                            BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                            BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                            BeStockRes.Uds_lic_plate = 0
                                            BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                            BeStockRes.No_bulto = vStockOrigen.No_bulto
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.IdPicking = 0
                                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                            BeStockRes.IdDespacho = 0
                                            BeStockRes.añada = vStockOrigen.Añada
                                            BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                            BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                            BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                            BeStockRes.Host = MaquinaQueSolicita
                                            BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                            CantidadStockDestino = BeStockRes.Cantidad

                                            clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), BeBodega.Permitir_Decimales)

                                            Insertar(BeStockRes,
                                                     lConnection,
                                                     ltransaction)

                                            vNombreCasoReservaInternoWMS = "CASO_#25_EJC202310090957"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            If Not pBeTrasladoDet Is Nothing Then

                                                If BeStockRes.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    If BePedidoDet.IdPresentacion = 0 Then
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                    Else
                                                        pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                    End If
                                                End If

                                                clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                                 BeProducto,
                                                                                                                 lConnection,
                                                                                                                 ltransaction)
                                            End If


                                            Restar_Stock_Reservado(lBeStockExistente,
                                                                       pBeConfigEnc,
                                                                       lConnection,
                                                                       ltransaction)

                                            lBeStockAReservar.Add(BeStockRes)


                                            lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                            FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                            DiasVencimiento,
                                                                                                            pBeConfigEnc,
                                                                                                            lConnection,
                                                                                                            ltransaction,
                                                                                                            BeProducto,
                                                                                                            pTarea_Reabasto,
                                                                                                            vFechaMinimaVenceZonaPicking,
                                                                                                            vFechaMinimaVenceZonaALM,
                                                                                                            lBeStockExistente,
                                                                                                            BePresentacionDefecto)

                                            vCantidadCompletada = (vCantidadPendiente = 0)

                                            If vCantidadCompletada Then
                                                Exit For
                                            End If

                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)
                                            vCantidadPendienteEnPres -= Math.Round(vCantidadDispStock * BePresentacionDefecto.Factor, 6)

                                        End If

                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                        BeStockRes.Lote = vStockOrigen.Lote
                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                        BeStockRes.Peso = vStockOrigen.Peso
                                        BeStockRes.Estado = "UNCOMMITED"
                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                        BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.IdPicking = 0
                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                        BeStockRes.IdDespacho = 0
                                        BeStockRes.añada = vStockOrigen.Añada
                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.Host = MaquinaQueSolicita
                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                        CantidadStockDestino = BeStockRes.Cantidad

                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                        Insertar(BeStockRes,
                                                 lConnection,
                                                 ltransaction)

                                        vNombreCasoReservaInternoWMS = "CASO_#26_EJC202310090957"
                                        vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                        If Not pBeTrasladoDet Is Nothing Then

                                            If BeStockRes.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                If BePedidoDet.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If
                                            End If

                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                         BeProducto,
                                                                                                         lConnection,
                                                                                                         ltransaction)
                                        End If


                                        Restar_Stock_Reservado(lBeStockExistente,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                        vCantidadCompletada = (vCantidadPendiente = 0)
                                        lBeStockAReservar.Add(BeStockRes)

                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                            DiasVencimiento,
                                                                                                            pBeConfigEnc,
                                                                                                            lConnection,
                                                                                                            ltransaction,
                                                                                                            BeProducto,
                                                                                                            pTarea_Reabasto,
                                                                                                            vFechaMinimaVenceZonaPicking,
                                                                                                            vFechaMinimaVenceZonaALM,
                                                                                                            lBeStockExistente,
                                                                                                            BePresentacionDefecto)

                                        If vCantidadCompletada Then
                                            Dim Log_20230301_N As String = String.Format("Log_202303011308N: Cantidad_Completada_202303011301L: Código: {0} Sol: {1} Reservado: {2}. " & vbNewLine,
                                                                                          BeProducto.Codigo,
                                                                                          vCantidadSolicitadaPedido,
                                                                                          BeStockRes.Cantidad)
                                            clsLnLog_error_wms.Agregar_Error(Log_20230301_N)
                                            Exit For
                                        End If

                                    Else

                                        If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                            BePresentacionDefecto = New clsBeProducto_Presentacion()
                                            BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                    lConnection,
                                                                                                                                    ltransaction)

                                            If Not BePresentacionDefecto Is Nothing Then

                                                vIndicePresentacion = lPresentaciones.FindIndex(Function(x) x.IdPresentacion = BePresentacionDefecto.IdPresentacion)

                                                If vIndicePresentacion <> -1 Then
                                                    BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                    BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                                Else
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If

                                                vSolicitudEsEnUMBas = True

                                            End If


                                        ElseIf vStockOrigen.IdPresentacion <> 0 Then

                                            If vIndicePresentacion <> -1 Then
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto = lPresentaciones(vIndicePresentacion).Clone()
                                            Else
                                                BePresentacionDefecto = New clsBeProducto_Presentacion()
                                                BePresentacionDefecto.IdPresentacion = vStockOrigen.IdPresentacion
                                                clsLnProducto_presentacion.GetSingle(BePresentacionDefecto, lConnection, ltransaction)
                                                If Not BePresentacionDefecto Is Nothing Then
                                                    lPresentaciones.Add(BePresentacionDefecto.Clone())
                                                End If
                                            End If

                                            vSolicitudEsEnUMBas = False

                                        End If

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadEnStockEnPres = Math.Round(vCantidadDispStock / BePresentacionDefecto.Factor, 6)
                                        End If


                                        Split_Decimal(vCantidadEnStockEnPres, vCantidadEnteraStockPres, vCantidadDecimalStockUMBas)
                                        Split_Decimal(pStockResSolicitud.Peso, vPesoEnteroPresStock, vPesoDecimalStockUMBas)

                                        If Not BePresentacionDefecto Is Nothing Then
                                            vCantidadDecimalUMBasStock = Math.Ceiling(Math.Round(vCantidadDecimalStockUMBas * BePresentacionDefecto.Factor, 2))
                                        Else

                                            vCantidadDecimalUMBasStock = vCantidadDecimalStockUMBas
                                        End If

                                        If (vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) Then

                                            vCantidadPendienteEnPres = vCantidadPendiente
                                            vCantidadDispStockEnPres = vStockOrigen.Cantidad

                                        Else

                                            vCantidadPendienteEnPres = vCantidadPendiente

                                            If pStockResSolicitud.IdPresentacion = 0 Then
                                                vCantidadDispStockEnPres = vStockOrigen.Cantidad
                                            Else
                                                vCantidadDispStockEnPres = Math.Round(vStockOrigen.Cantidad / BePresentacionDefecto.Factor, 6)
                                            End If

                                            If Not (vStockOrigen.IdPresentacion = 0) Then
                                                If Not vConvirtioCantidadSolicitadaEnUmBas Then
                                                    vCantidadPendiente = Math.Round(vCantidadPendiente * BePresentacionDefecto.Factor, 6)
                                                    vConvirtioCantidadSolicitadaEnUmBas = True
                                                End If
                                            Else
                                                vCantidadPendiente = vCantidadPendiente
                                            End If

                                        End If

                                        If vSolicitudEsEnUMBas Then
                                            If vCantidadPendienteEnPres > vCantidadDispStockEnPres Then
                                                If Not ((vStockOrigen.IdPresentacion = 0 AndAlso pStockResSolicitud.IdPresentacion = 0) AndAlso vCantidadDispStockEnPres < BePresentacionDefecto.Factor) Then
                                                    clsLnLog_error_wms.Agregar_Error("#EJC202302081729: Condición para reserva de unidades alcanzada, impacto desconocido.")
                                                End If
                                            End If
                                        End If

                                        BeStockRes = New clsBeStock_res
                                        BeStockRes.IdTransaccion = pStockResSolicitud.IdTransaccion
                                        BeStockRes.Indicador = IIf(pStockResSolicitud.Indicador = "", "PED", pStockResSolicitud.Indicador)
                                        BeStockRes.IdBodega = vStockOrigen.IdBodega
                                        BeStockRes.IdStock = vStockOrigen.IdStock
                                        BeStockRes.IdPropietarioBodega = vStockOrigen.IdPropietarioBodega
                                        BeStockRes.IdProductoBodega = vStockOrigen.IdProductoBodega

                                        If Not vSolicitudEsEnUMBas Then
                                            BeStockRes.IdPresentacion = IIf(pStockResSolicitud.IdPresentacion = 0, 0, vStockOrigen.Presentacion.IdPresentacion)
                                        End If

                                        If vCantidadPendiente = vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadDispStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        ElseIf vCantidadPendiente < vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadPendiente
                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        ElseIf vCantidadPendiente > vCantidadDispStock Then

                                            vCantidadAReservarPorIdStock = vCantidadDispStock
                                            vCantidadPendiente -= vCantidadAReservarPorIdStock
                                            vCantidadPendiente = Math.Round(vCantidadPendiente, 6)

                                        End If

                                        BeStockRes.IdUbicacion = vStockOrigen.IdUbicacion
                                        BeStockRes.IdProductoEstado = vStockOrigen.ProductoEstado.IdEstado
                                        BeStockRes.IdUnidadMedida = vStockOrigen.IdUnidadMedida
                                        BeStockRes.Lote = vStockOrigen.Lote
                                        BeStockRes.Lic_plate = vStockOrigen.Lic_plate
                                        BeStockRes.Serial = IIf(No_Linea <> 0, No_Linea, vStockOrigen.Serial)
                                        BeStockRes.Peso = vStockOrigen.Peso
                                        BeStockRes.Estado = "UNCOMMITED"
                                        BeStockRes.Fecha_ingreso = vStockOrigen.Fecha_Ingreso
                                        BeStockRes.Fecha_vence = vStockOrigen.Fecha_vence
                                        BeStockRes.Uds_lic_plate = vStockOrigen.Uds_lic_plate
                                        BeStockRes.Ubicacion_ant = vStockOrigen.IdUbicacion_anterior
                                        BeStockRes.No_bulto = vStockOrigen.No_bulto
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.IdPicking = 0
                                        BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                                        BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                                        BeStockRes.IdDespacho = 0
                                        BeStockRes.añada = vStockOrigen.Añada
                                        BeStockRes.Fecha_manufactura = vStockOrigen.Fecha_Manufactura
                                        BeStockRes.Cantidad = Math.Round(vCantidadAReservarPorIdStock, 6)
                                        BeStockRes.IdRecepcion = vStockOrigen.IdRecepcionEnc
                                        BeStockRes.Host = MaquinaQueSolicita
                                        BeStockRes.IdStockRes = MaxID(lConnection, ltransaction) + 1

                                        CantidadStockDestino = BeStockRes.Cantidad

                                        vPermitirDecimales = clsLnBodega.Get_Permitir_Decimales(BeStockRes.IdBodega, lConnection, ltransaction)
                                        clsPublic.Abs(CantidadStockDestino - Fix(CantidadStockDestino), vPermitirDecimales)

                                        Insertar(BeStockRes,
                                                 lConnection,
                                                 ltransaction)

                                        vNombreCasoReservaInternoWMS = "CASO_#27_EJC202310090957"
                                        vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                        If Not pBeTrasladoDet Is Nothing Then

                                            If BeStockRes.IdPresentacion = 0 Then
                                                pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                            Else
                                                If BePedidoDet.IdPresentacion = 0 Then
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += BeStockRes.Cantidad
                                                Else
                                                    pBeTrasladoDet.Quantity_Reserved_WMS += Math.Round(BeStockRes.Cantidad / BePresentacionDefecto.Factor, 6)
                                                End If
                                            End If

                                            clsLnI_nav_ped_traslado_det.Actualizar_Quantity_Reserved_WMS(pBeTrasladoDet,
                                                                                                         BeProducto,
                                                                                                         lConnection,
                                                                                                         ltransaction)
                                        End If


                                        Restar_Stock_Reservado(lBeStockExistente,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                        vCantidadCompletada = (vCantidadPendiente = 0)
                                        lBeStockAReservar.Add(BeStockRes)

                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                        '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3
                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                         DiasVencimiento,
                                                                                                         pBeConfigEnc,
                                                                                                         lConnection,
                                                                                                         ltransaction,
                                                                                                         BeProducto,
                                                                                                         pTarea_Reabasto,
                                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                                         vFechaMinimaVenceZonaALM,
                                                                                                         lBeStockExistente,
                                                                                                         BePresentacionDefecto)

                                        If vCantidadCompletada Then
                                            Exit For
                                        End If

                                    End If

                                End If

                            Next

                        End If

                    End If

                    If lBeStockAReservar.Count > 0 Then

                        If Inserta_Stock_Reservado(lBeStockAReservar,
                                                   lConnection,
                                                   ltransaction) Then

                            pListStockResOUT.AddRange(lBeStockAReservar)

                            If Not (vCantidadCompletada AndAlso vCantidadPendiente > 0) AndAlso (pStockResSolicitud.IdPresentacion = 0 AndAlso vCantidadDecimalUMBas = 0) Then
                                vCantidadDecimalUMBas = vCantidadPendiente
                            ElseIf Not ((vCantidadCompletada AndAlso vCantidadPendiente > 0) AndAlso (pStockResSolicitud.IdPresentacion <> 0 AndAlso vCantidadDecimalUMBas = 0 AndAlso pBeConfigEnc.Explosion_Automatica)) AndAlso vBusquedaEnUmBas Then
                                If vCantidadDecimalUMBas > 0 Then
                                    vCantidadPendiente = vCantidadDecimalUMBas
                                Else
                                    vCantidadDecimalUMBas = vCantidadPendiente
                                End If
                            End If

                            Reserva_Stock_From_MI3 = True

                            If (pBeConfigEnc.Explosion_Automatica) AndAlso ((vCantidadDecimalUMBas > 0) OrElse (vCantidadPendiente > 0)) Then

                                If (vCantidadDecimalUMBas > 0) OrElse (vCantidadPendiente > 0) Then

                                    Dim BeStockResUMBas As New clsBeStock_res
                                    BeStockResUMBas = BeStockRes.Clone()

                                    If vCantidadPendiente > 0 AndAlso Not (vCantidadDecimalUMBas = vCantidadPendiente) Then
                                        vCantidadDecimalUMBas += vCantidadPendiente
                                    End If

                                    BeStockResUMBas.Cantidad = vCantidadDecimalUMBas
                                    BeStockResUMBas.IdPresentacion = 0
                                    BeStockResUMBas.Serial = No_Linea

                                    If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                        BeStockResUMBas.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                    End If

                                    Dim vCantDisRef As Double = 0

                                    Dim ExcluirUbicacionesPicking As Boolean

                                    ExcluirUbicacionesPicking = pTarea_Reabasto

                                    If lBeStockExistente.Count = 0 Then

                                        If pStockResSolicitud.IdPresentacion = 0 Then vBusquedaEnUmBas = True

                                        lBeStockExistente = clsLnStock.lStock(pStockResSolicitud,
                                                                              BeProducto,
                                                                              DiasVencimiento,
                                                                              pBeConfigEnc,
                                                                              lConnection,
                                                                              ltransaction,
                                                                              True,
                                                                              True,
                                                                              pTarea_Reabasto,
                                                                              pEs_Devolucion)

                                        Restar_Stock_Reservado(lBeStockExistente,
                                                               pBeConfigEnc,
                                                               lConnection,
                                                               ltransaction)

                                        vRestoInventarioEnUmBas = True

                                        lBeStockExistente = lBeStockExistente.Where(Function(x) x.Cantidad > 0).ToList()

                                        '#EJC20231019_Get_Fecha_Vence_Minima_Stock_Reserva_MI3
                                        FechaMinimaVenceStock = Get_Fecha_Vence_Minima_Stock_Reserva_MI3(pStockResSolicitud,
                                                                                                         DiasVencimiento,
                                                                                                         pBeConfigEnc,
                                                                                                         lConnection,
                                                                                                         ltransaction,
                                                                                                         BeProducto,
                                                                                                         pTarea_Reabasto,
                                                                                                         vFechaMinimaVenceZonaPicking,
                                                                                                         vFechaMinimaVenceZonaALM,
                                                                                                         lBeStockExistente,
                                                                                                         BePresentacionDefecto)

                                        If pStockResSolicitud.IdPresentacion = 0 Then
                                            vZonaNoPickingStockEnUmBas = lBeStockExistente.Count > 0
                                        Else
                                            Debug.WriteLine("vZonaNoPickingStockEnUmBas = false")
                                        End If

                                    End If

                                    If pStockResSolicitud.IdPresentacion = 0 And Not pTarea_Reabasto And vZonaNoPickingStockEnUmBas Then
                                        ExcluirUbicacionesPicking = True
                                    End If

                                    'No tengo unidades en almacenaje.
                                    If vZonaNoPickingStockEnUmBas Then

                                        If Not vNombreCasoReservaInternoWMS.Contains("_LLR_CASO_#28_") Then

                                            vNombreCasoReservaInternoWMS += "_LLR_CASO_#28_"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            '#CKFK20250426 asignarle a la nueva solicitud de inventario el idubicacionAbastecerCon
                                            If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                                BeStockResUMBas.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                            End If

                                            '#CKFK20240116 Modifiqué el enviar ExcluirUbicacionesPicking por pTareaReabasto porque siempre se debe enviar en falso
                                            If Not Reserva_Stock_From_MI3(BeStockResUMBas,
                                                                          DiasVencimiento,
                                                                          MaquinaQueSolicita,
                                                                          pBeConfigEnc,
                                                                          vCantDisRef,
                                                                          pIdPropietarioBodega,
                                                                          vlBeStockAReservarUMBas,
                                                                          lConnection,
                                                                          ltransaction,
                                                                          No_Linea,
                                                                          pTarea_Reabasto,
                                                                          pBeTrasladoDet) Then

                                                If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                    pCantidadDisponibleStock = vCantidadDispStock
                                                    Throw New Exception(String.Format("{0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                              BeProducto.Codigo,
                                                                              vCantidadSolicitadaPedido,
                                                                              vCantidadStock))

                                                Else

                                                    If vlBeStockAReservarUMBas.Count > 0 Then
                                                        Dim vlBeStockAReservar As New List(Of clsBeStock_res)
                                                        vlBeStockAReservar = lBeStockAReservar
                                                        pListStockResOUT.AddRange(vlBeStockAReservar)
                                                    End If

                                                    'Reserva_Stock_From_MI3 = False

                                                End If
                                            Else

                                                '#CKFK20240808 Agregué estas variables para que se sepa que si se pudo completar la reserva
                                                vCantidadPendiente = 0
                                                vCantidadCompletada = True
                                                Reserva_Stock_From_MI3 = True
                                                clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, lConnection, ltransaction)

                                            End If

                                        End If

                                    Else

                                        If Not vNombreCasoReservaInternoWMS.Contains("LLR_CASO_#29") Then

                                            vNombreCasoReservaInternoWMS += "_LLR_CASO_#29_"
                                            vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                            '#EJC20231031: Si la solicitud es presentación, mantener la presentación en la llamada recursiva.
                                            If Not (pStockResSolicitud.IdPresentacion = 0) AndAlso Not vBusquedaEnUmBas Then

                                                If vCantidadPendiente = 0 AndAlso vCantidadDecimalUMBasStock > 0 Then
                                                    BeStockResUMBas.IdPresentacion = 0
                                                Else
                                                    BeStockResUMBas.IdPresentacion = pStockResSolicitud.IdPresentacion
                                                End If

                                                If vCantidadPendiente > 0 Then

                                                    If BePresentacionDefecto.IdPresentacion = 0 Then
                                                        BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                                                  lConnection,
                                                                                                                                                  ltransaction)
                                                    End If

                                                    vCantidadPendiente = Math.Round(vCantidadPendiente / BePresentacionDefecto.Factor, 6)
                                                    BeStockResUMBas.Cantidad = vCantidadPendiente

                                                End If

                                            End If

                                            '#CKFK20250426 asignarle a la nueva solicitud de inventario el idubicacionAbastecerCon
                                            If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                                BeStockResUMBas.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                            End If

                                            If Not Reserva_Stock_From_MI3(BeStockResUMBas,
                                                                          DiasVencimiento,
                                                                          MaquinaQueSolicita,
                                                                          pBeConfigEnc,
                                                                          vCantDisRef,
                                                                          pIdPropietarioBodega,
                                                                          vlBeStockAReservarUMBas,
                                                                          lConnection,
                                                                          ltransaction,
                                                                          No_Linea,
                                                                          False,
                                                                          pBeTrasladoDet) Then

                                                If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                    pCantidadDisponibleStock = vCantidadDispStock
                                                    Throw New Exception(String.Format("{0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                  BeProducto.Codigo,
                                                                                  vCantidadSolicitadaPedido,
                                                                                  vCantidadStock))
                                                Else

                                                    '#EJC202312151817:Validar si esto aplica o no.
                                                    If vlBeStockAReservarUMBas.Count > 0 Then
                                                        Dim vlBeStockAReservar As New List(Of clsBeStock_res)
                                                        vlBeStockAReservar = lBeStockAReservar
                                                        pListStockResOUT.AddRange(vlBeStockAReservar)
                                                    End If

                                                    'Exit Function

                                                End If

                                            Else
                                                '#CKFK20240212 Agregué estas dos asginaciones vCantidadPendiente = 0 vCantidadCompletada = True
                                                vCantidadPendiente = 0
                                                vCantidadCompletada = True
                                                Reserva_Stock_From_MI3 = True
                                                clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, lConnection, ltransaction)
                                            End If

                                        Else

                                            If Not vNombreCasoReservaInternoWMS.Contains("LLR_CASO_#30_") Then

                                                vNombreCasoReservaInternoWMS += "_LLR_CASO_#30_"
                                                vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                                '#CKFK20240129 Preguntar a Erik si aquí no va el BeStockResUMBas
                                                clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                                BeStockResUMBas.IdPresentacion = 0
                                                BeStockResUMBas.Cantidad = vCantidadPendiente

                                                '#CKFK20250426 asignarle a la nueva solicitud de inventario el idubicacionAbastecerCon
                                                If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                                    BeStockResUMBas.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                                End If

                                                If Not Reserva_Stock_From_MI3(BeStockResUMBas,
                                                                              DiasVencimiento,
                                                                              MaquinaQueSolicita,
                                                                              pBeConfigEnc,
                                                                              vCantDisRef,
                                                                              pIdPropietarioBodega,
                                                                              vlBeStockAReservarUMBas,
                                                                              lConnection,
                                                                              ltransaction,
                                                                              No_Linea,
                                                                              False,
                                                                              pBeTrasladoDet) Then

                                                    If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                        pCantidadDisponibleStock = vCantidadDispStock
                                                        Throw New Exception(String.Format("{0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                      BeProducto.Codigo,
                                                                                      vCantidadSolicitadaPedido,
                                                                                      vCantidadStock))

                                                    Else

                                                        '#EJC202312151817:Validar si esto aplica o no.
                                                        If vlBeStockAReservarUMBas.Count > 0 Then
                                                            pListStockResOUT.AddRange(lBeStockAReservar)
                                                        End If

                                                        'Exit Function

                                                    End If

                                                Else

                                                    '#CKFK20240808 Agregué estas variables para que se sepa que si se pudo completar la reserva
                                                    vCantidadPendiente = 0
                                                    vCantidadCompletada = True
                                                    Reserva_Stock_From_MI3 = True
                                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, lConnection, ltransaction)

                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            ElseIf Not vCantidadCompletada Then

                                '#CKFK20240128 Agregué este llamado recursivo por error en Mercopan
                                If Not vNombreCasoReservaInternoWMS.Contains("LLR_CASO_#31_") Then

                                    vNombreCasoReservaInternoWMS += "_LLR_CASO_#31_"
                                    vMensajeReserva = vNombreCasoReservaInternoWMS &
                                                                            " Fecha Mínima: " & FechaMinimaVenceStock.Date &
                                                                            " DiasVencimiento: " & DiasVencimiento & " FechaMinimaVenceZonaPicking: " & vFechaMinimaVenceZonaPicking.Date &
                                                                            " vFechaMinimaVenceZonaALM: " & vFechaMinimaVenceZonaALM.Date &
                                                                            " FechaReservada: " & BeStockRes.Fecha_vence.Date &
                                                                            " Lote: " & BeStockRes.Lote &
                                                                            " Ubicación: " & BeStockRes.IdUbicacion

                                    clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeReserva)

                                    Dim vCantDisRef As Double = 0

                                    Dim BeStockResPresFalt As New clsBeStock_res
                                    BeStockResPresFalt = BeStockRes.Clone()

                                    If Not BePresentacionDefecto Is Nothing Then
                                        If Not pStockResSolicitud.IdPresentacion = 0 Then
                                            BeStockResPresFalt.Cantidad = vCantidadPendiente / BePresentacionDefecto.Factor
                                        Else
                                            BeStockResPresFalt.Cantidad = vCantidadPendiente
                                        End If
                                    Else
                                        BeStockResPresFalt.Cantidad = vCantidadPendiente
                                    End If

                                    '#CKFK20250426 Agregar al nuevo stock a reservar el idubicacionabastecercon
                                    If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                        BeStockResPresFalt.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                    End If

                                    If Not Reserva_Stock_From_MI3(BeStockResPresFalt,
                                                                  DiasVencimiento,
                                                                  MaquinaQueSolicita,
                                                                  pBeConfigEnc,
                                                                  vCantDisRef,
                                                                  pIdPropietarioBodega,
                                                                  vlBeStockAReservarPresFaltante,
                                                                  lConnection,
                                                                  ltransaction,
                                                                  No_Linea,
                                                                  False,
                                                                  pBeTrasladoDet) Then

                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                            pCantidadDisponibleStock = vCantidadDispStock
                                            Throw New Exception(String.Format("{0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                      BeProducto.Codigo,
                                                                                      vCantidadSolicitadaPedido,
                                                                                      vCantidadStock))

                                        Else

                                            '#EJC202312151817:Validar si esto aplica o no.
                                            If vlBeStockAReservarPresFaltante.Count > 0 Then
                                                pListStockResOUT.AddRange(vlBeStockAReservarPresFaltante)
                                            End If

                                            'Exit Function

                                        End If

                                    Else
                                        '#CKFK20240808 Agregué estas variables para que se sepa que si se pudo completar la reserva
                                        vCantidadPendiente = 0
                                        vCantidadCompletada = True
                                        Reserva_Stock_From_MI3 = True
                                        clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, lConnection, ltransaction)
                                    End If

                                End If


                            End If

                            pStockResSolicitud = BeStockRes

                            If Not vlBeStockAReservarUMBas Is Nothing Then
                                If vlBeStockAReservarUMBas.Count > 0 Then
                                    Dim vNlBeStockAReservarUMBas As New List(Of clsBeStock_res)
                                    vNlBeStockAReservarUMBas = vlBeStockAReservarUMBas
                                    lBeStockAReservar.AddRange(vNlBeStockAReservarUMBas)
                                ElseIf Not vlBeStockAReservarPresFaltante Is Nothing Then
                                    If vlBeStockAReservarPresFaltante.Count > 0 Then
                                        Dim vNlBeStockAReservarPresFaltante As New List(Of clsBeStock_res)
                                        vNlBeStockAReservarPresFaltante = vlBeStockAReservarPresFaltante
                                        lBeStockAReservar.AddRange(vNlBeStockAReservarPresFaltante)
                                    End If
                                End If
                            End If

                            pListStockResOUT = lBeStockAReservar

                            '#CKFK20240129 Agregué esta condición para que no se vaya de la función sin determinar correctamente
                            'si la reserva está completa o no, anteriormente solo se colocaba el resultado en true Reserva_Stock_From_MI3 = True

                            Reserva_Stock_From_MI3 = vCantidadCompletada

                            If Not vCantidadCompletada Then

                                If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                    vMensajeNoExplosionEnZonasNoPicking = String.Format("{0} IdProductoBodega: {2} Producto Solicitado: {1} UM: {3} Presentacion {4} Cantidad: {5}",
                                                                                        clsDalEx.ErrorS0005,
                                                                                        BeProducto.Codigo,
                                                                                        pStockResSolicitud.IdProductoBodega,
                                                                                        BeProducto.UnidadMedida.Nombre,
                                                                                        pStockResSolicitud.IdPresentacion,
                                                                                        pStockResSolicitud.Cantidad)

                                    If Not pBeTrasladoDet Is Nothing Then
                                        '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                        pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                        pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                        clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                              lConnection,
                                                                                              ltransaction)
                                    End If

                                    Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                Else

                                    vMensajeNoExplosionEnZonasNoPicking = "#MI3_2312201922: No se pudo reservar el stock y la bandera rechazar_pedido_incompleto = No."

                                    If Not pBeTrasladoDet Is Nothing Then
                                        '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                        pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                        pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                        clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                              lConnection,
                                                                                              ltransaction)
                                    End If

                                    Reserva_Stock_From_MI3 = False

                                End If

                            End If

                        Else
                            If Not vOrdernarListaStockSinPresentacionPrimero Then
                                Throw New Exception(String.Format("{0} IdProductoBodega: {2} Producto Solicitado: {1}", clsDalEx.ErrorS0006, BeProducto.Codigo, pStockResSolicitud.IdProductoBodega))
                            Else
                                Throw New Exception(vbNewLine & String.Format("{0} IdProductoBodega: {2} Producto Solicitado: {1}", clsDalEx.ErrorS0002, BeProducto.Codigo, pStockResSolicitud.IdProductoBodega))
                            End If
                        End If

                    Else
                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                            vMensajeNoExplosionEnZonasNoPicking = String.Format("{0} IdProductoBodega: {2} Producto Solicitado: {1}UM: {3} Presentacion {4} Cantidad: {5}",
                                                                                        clsDalEx.ErrorS0005,
                                                                                        BeProducto.Codigo,
                                                                                        pStockResSolicitud.IdProductoBodega,
                                                                                        BeProducto.UnidadMedida.Nombre,
                                                                                        pStockResSolicitud.IdPresentacion,
                                                                                        pStockResSolicitud.Cantidad)
                            '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                            pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                            pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                            clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                  lConnection,
                                                                                  ltransaction)

                            Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                        Else
                            vMensajeNoExplosionEnZonasNoPicking = "#MI3_2312201922: No se pudo explosionar en zonas de almacenamiento (Rack), la bandera rechazar_pedido_incompleto = No."

                            BeStockRes.IdPedido = pStockResSolicitud.IdPedido
                            BeStockRes.IdPedidoDet = pStockResSolicitud.IdPedidoDet
                            BeStockRes.Cantidad = pStockResSolicitud.Cantidad

                            vNombreCasoReservaInternoWMS = "#SR240315"

                            clsLnTrans_pe_det_log_reserva.Agregar_Log_Reserva(BeStockRes, vNombreCasoReservaInternoWMS, vMensajeNoExplosionEnZonasNoPicking)

                            '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                            If Not pBeTrasladoDet Is Nothing Then
                                pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                      lConnection,
                                                                                      ltransaction)
                            End If

                            Reserva_Stock_From_MI3 = False
                            Exit Function
                        End If
                    End If

                Else

                    If pStockResSolicitud.IdPresentacion <> 0 Then

                        If (pBeConfigEnc.Explosion_Automatica) Then

                            Split_Decimal(pStockResSolicitud.Cantidad,
                                              vCantidadEnteraPres,
                                              vCantidadDecimalUMBas)

                            If IdProducto = 0 Then
                                IdProducto = clsLnProducto_bodega.Get_IdProducto_By_IdProductoBodega(pStockResSolicitud.IdProductoBodega,
                                                                                                     lConnection,
                                                                                                     ltransaction)
                            End If

                            If IdProducto <> 0 Then
                                BePresentacionDefecto = clsLnProducto_presentacion.Get_Presentacion_Defecto_By_IdProducto(IdProducto,
                                                                                                                        lConnection,
                                                                                                                        ltransaction)

                            End If

                            If Not BePresentacionDefecto Is Nothing Then

                                If Not BePresentacionDefecto.Factor = 0 Then

                                    vCantidadDecimalUMBas = Math.Ceiling(Math.Round(vCantidadDecimalUMBas * BePresentacionDefecto.Factor, 2))

                                    If vCantidadEnteraPres > 0 Then
                                        vCantidadSolicitadaPedido = vCantidadEnteraPres
                                    Else
                                        vCantidadSolicitadaPedido = vCantidadDecimalUMBas
                                    End If

                                End If

                            End If

                        Else
                            vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad * BePresentacionDefecto.Factor
                            vConvirtioCantidadSolicitadaEnUmBas = True
                        End If

                        If vCantidadDecimalUMBas > 0 Then

                            If (pBeConfigEnc.Explosion_Automatica) AndAlso (vCantidadDecimalUMBas > 0) Then

                                Dim BeStockResUMBas As New clsBeStock_res
                                BeStockResUMBas = pStockResSolicitud.Clone()

                                If BeStockResUMBas.No_bulto = 0 Then

                                    BeStockResUMBas.Cantidad = vCantidadDecimalUMBas
                                    BeStockResUMBas.Atributo_Variante_1 = Nothing
                                    BeStockResUMBas.IdPresentacion = 0
                                    BeStockResUMBas.Serial = No_Linea
                                    BeStockResUMBas.No_bulto = 1965 'Identificación para solicitud recursiva.

                                    If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                        BeStockResUMBas.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                    End If

                                    Dim vCantDisRef As Double = 0

                                    If Inserta_Stock_Reservado(lBeStockAReservar, lConnection, ltransaction) Then

                                        If Not Reserva_Stock_From_MI3(BeStockResUMBas,
                                                                          DiasVencimiento,
                                                                          MaquinaQueSolicita,
                                                                          pBeConfigEnc,
                                                                          vCantDisRef,
                                                                          pIdPropietarioBodega,
                                                                          vlBeStockAReservarUMBas,
                                                                          lConnection,
                                                                          ltransaction,
                                                                          No_Linea,
                                                                          False,
                                                                          pBeTrasladoDet) Then

                                            If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then

                                                pCantidadDisponibleStock = vCantidadDispStock

                                                vMensajeNoExplosionEnZonasNoPicking = String.Format("{0} Código: {1} Sol: {2} Disp: {3}. " & vbNewLine, clsDalEx.ErrorS0002,
                                                                                                      BeProducto.Codigo,
                                                                                                      vCantidadSolicitadaPedido,
                                                                                                      vCantidadStock)
                                                '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                                pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                                pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                                clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                                      lConnection,
                                                                                                      ltransaction)

                                                Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                            Else
                                                Exit Function
                                            End If

                                        End If

                                        '#CKFK20230207 Agregué esto porque no se agrega a la respuesta
                                        pStockResSolicitud = BeStockResUMBas

                                        If Not vlBeStockAReservarUMBas Is Nothing Then
                                            If vlBeStockAReservarUMBas.Count > 0 Then
                                                lBeStockAReservar.AddRange(vlBeStockAReservarUMBas)
                                            End If
                                        End If

                                        pListStockResOUT = lBeStockAReservar
                                        Reserva_Stock_From_MI3 = True

                                    End If

                                    '#EJC20210707: Igualar el stock reservador con la solicitud.
                                    pStockResSolicitud = BeStockRes

                                    If Not vlBeStockAReservarUMBas Is Nothing Then
                                        If vlBeStockAReservarUMBas.Count > 0 Then
                                            lBeStockAReservar.AddRange(vlBeStockAReservarUMBas)
                                        End If
                                    End If

                                    '#EJC20220627_0927:Devolver la lista de stock reservado para adicionar en picking existente.
                                    pListStockResOUT = lBeStockAReservar
                                    Reserva_Stock_From_MI3 = True

                                Else
                                    Exit Function
                                End If

                            End If

                        Else

                            If pBeConfigEnc.Despachar_existencia_parcial = tDespacharExistenciaParcial.No Then

                                If pStockResSolicitud.IdPresentacion = 0 Then


                                    vMensajeNoExplosionEnZonasNoPicking = String.Format("{0} Código:{1} UMBas={2} Pres='Sin pres' Cant={3} (Verifique si tiene existencia en UMBas en WMS) - Config_IF: Despachar_existencia_parcial = No ", clsDalEx.ErrorS0004,
                                                                                                                                                                    BeProducto.Codigo,
                                                                                                                                                                    BeProducto.UnidadMedida.Nombre,
                                                                                                                                                                    pStockResSolicitud.Cantidad)
                                    '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                    pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                    pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                    clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                          lConnection,
                                                                                          ltransaction)

                                    Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                Else

                                    vMensajeNoExplosionEnZonasNoPicking = String.Format("{0} Código:{1} UMBas={2} Pres={3} Cant={4} (Verifique si tiene existencia en Pres. en WMS) - Config_IF: Despachar_existencia_parcial = No ", clsDalEx.ErrorS0004,
                                                                                                                                                             BeProducto.Codigo,
                                                                                                                                                             BeProducto.UnidadMedida.Nombre,
                                                                                                                                                             pStockResSolicitud.IdPresentacion,
                                                                                                                                                             pStockResSolicitud.Cantidad)
                                    '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                    pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                    pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                    clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                          lConnection,
                                                                                          ltransaction)

                                    Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                End If

                            End If

                        End If

                    Else

                        'IDENTIFICACIÓN DE RECURSIVIDAD.
                        '#CKFK20230202 No llega la cantidad pendiente solo se sabe que no está completo
                        vCantidadSolicitadaPedido = pStockResSolicitud.Cantidad

                        '#CKFK20230202 Aquí necesitamos mandar a explosionar el producto
                        If (pBeConfigEnc.Explosion_Automatica) AndAlso (vCantidadSolicitadaPedido > 0) AndAlso lBeStockExistente.Count > 0 Then

                            Dim BeStockResUMBas As New clsBeStock_res
                            BeStockResUMBas = pStockResSolicitud.Clone()

                            If BeStockResUMBas.No_bulto = 0 Then

                                BeStockResUMBas.Cantidad = vCantidadSolicitadaPedido
                                BeStockResUMBas.IdPresentacion = 0
                                BeStockResUMBas.Serial = No_Linea
                                BeStockResUMBas.No_bulto = 1965 'Identificación para solicitud recursiva.

                                If pStockResSolicitud.IdUbicacionAbastecerCon <> 0 Then
                                    BeStockResUMBas.IdUbicacionAbastecerCon = pStockResSolicitud.IdUbicacionAbastecerCon
                                End If

                                Dim vCantDisRef As Double = 0

                                If Inserta_Stock_Reservado(lBeStockAReservar,
                                                           lConnection,
                                                           ltransaction) Then

                                    If Not Reserva_Stock_From_MI3(BeStockResUMBas,
                                                                  DiasVencimiento,
                                                                  MaquinaQueSolicita,
                                                                  pBeConfigEnc,
                                                                  vCantDisRef,
                                                                  pIdPropietarioBodega,
                                                                  vlBeStockAReservarUMBas,
                                                                  lConnection,
                                                                  ltransaction,
                                                                  No_Linea,
                                                                  False,
                                                                  pBeTrasladoDet) Then

                                        pCantidadDisponibleStock = vCantidadDispStock

                                        vMensajeNoExplosionEnZonasNoPicking = vMensajeNoExplosionEnZonasNoPicking = String.Format("{0} Código:{1} UMBas={2} Pres='Sin pres' Cant={3} (Verifique si tiene existencia en UMBas en WMS) ", clsDalEx.ErrorS0004,
                                                                                                                                                                    BeProducto.Codigo,
                                                                                                                                                                    BeProducto.UnidadMedida.Nombre,
                                                                                                                                                                    pStockResSolicitud.Cantidad)
                                        '#EJC202401291004: Mejorar el mensaje cuando lleguen a este punto mis amados maestros.
                                        pBeTrasladoDet.Process_Result += vMensajeNoExplosionEnZonasNoPicking
                                        pBeTrasladoDet.Qty_to_Receive = vCantidadPendiente
                                        clsLnI_nav_ped_traslado_det.Actualizar_Process_Result(pBeTrasladoDet,
                                                                                            lConnection,
                                                                                            ltransaction)


                                        If pBeConfigEnc.Rechazar_pedido_incompleto = tRechazarPedidoIncompleto.Si Then
                                            Throw New Exception(vMensajeNoExplosionEnZonasNoPicking)
                                        Else
                                            Exit Function
                                        End If

                                    End If

                                End If

                                '#EJC20210707: Igualar el stock reservador con la solicitud.
                                pStockResSolicitud = BeStockRes

                                If Not vlBeStockAReservarUMBas Is Nothing Then
                                    If vlBeStockAReservarUMBas.Count > 0 Then
                                        lBeStockAReservar.AddRange(vlBeStockAReservarUMBas)
                                    End If
                                End If

                                '#EJC20220627_0927:Devolver la lista de stock reservado para adicionar en picking existente.
                                pListStockResOUT = lBeStockAReservar

                            Else
                                Exit Function
                            End If

                        End If

                    End If

                End If

            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function
