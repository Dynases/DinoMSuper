USE [DBDinoM]
GO
/****** Object:  Trigger [dbo].[Tr_Mam_UpdateInsert_Venta_Update]    Script Date: 15/12/2019 05:41:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER TRIGGER [dbo].[Tr_Mam_UpdateInsert_Venta_Update] ON [dbo].[TV0011]
AFTER UPDATE
AS
BEGIN

Declare 
		@tbnumi int,@tbtv1numi int, @tbty5prod int,@tbcmin decimal(18,2),@tbumin int, 
		@ingreso int, @salida int,@obs nvarchar(100),@cantAct decimal(18,2)
		,@maxid1 int,@fact date,@hact nvarchar(5),@uact nvarchar(10),@lcfpag date,@maxid2 int
		,@cantE decimal(18,2),@can decimal(18,2),@deposito int,@cliente nvarchar(100)
		,@tblote nvarchar(50),@tbfechavenc date
			set @salida = 4
					set @lcfpag=GETDATE ()


--Declarando el cursor
declare MiCursor Cursor
	for Select a.tbnumi ,a.tbtv1numi ,a.tbty5prod ,a.tbcmin ,a.tbumin,a.tblote ,a.tbfechaVenc    --, a.chhact, a.chuact, b.cpmov, b.cpdesc
				From inserted a where a.tbty5prod  >0 
--Abrir el cursor
open MiCursor
-- Navegar
Fetch MiCursor into @tbnumi,@tbtv1numi,@tbty5prod,@tbcmin,@tbumin,@tblote,@tbfechavenc
while (@@FETCH_STATUS = 0)
begin
			set @deposito =(Select b.abnumi   from TV001  as a,TA002 as b,TA001 as c where c.aata2depVenta  =b.abnumi 
and a.tanumi  =@tbtv1numi and a.taalm  =c.aanumi )

set @cliente =(select b.yddesc  from TV001 as a inner join TY004 as b on a.taclpr  =b.ydnumi and b.ydtip =1 and a.tanumi  =@tbtv1numi   )
	set @obs = CONCAT(' M ', ' - Venta numiprod:',@tbty5prod,'|',@cliente ,
	IIF((select top 1 Verlote from SY000)=1,Concat('   LOTE: ',@tblote,'   FECHA VENC:',@tbfechavenc),''))  --M=Modificar
		set @obs = CONCAT(@tbtv1numi,'-',@obs)
	--if(@fec >='2017/04/02' and @fact >='2017/04/02')
	--begin
		set @cantE = (select d.tbcmin   from deleted d where d.tbnumi   = @tbnumi )

			if (exists (select TI001.iccprod from TI001 where TI001.iccprod = convert(int, @tbty5prod)
			and TI001 .icalm =@deposito and TI001.iclot =@tblote and TI001.icfven =@tbfechavenc ))
			begin 	
				begin try
					begin tran Tr_UpdateTI001
						--Obtener la cantidad actual
						set @cantAct = (select TI001.iccven  from TI001 where TI001.iccprod  = convert(int, @tbty5prod)
						                                          and TI001 .icalm =@deposito 
																  and TI001.iclot =@tblote and TI001.icfven =@tbfechavenc )
						set @can = (@cantAct +(@cantE -@tbcmin ))

						--Actualizar Saldo Inventario
						update TI001 
							set iccven  = @can
							where TI001.iccprod  = CONVERT(int, @tbty5prod)  and TI001.icalm =@deposito 
			                   and TI001.iclot =@tblote and TI001.icfven =@tbfechavenc 
						--Modificar Movimiento
						--Cabecera
						set @fact=(SElect a.cafact   from TC001 as a where a.canumi  =@tbtv1numi )
						set @hact =(SElect a.cahact    from TC001 as a where a.canumi =@tbtv1numi )
						set @uact =(SElect a.cauact    from TC001 as a where a.canumi =@tbtv1numi )
						Update TI002 
							set ibconcep = @salida , ibobs = @obs, ibfact = @fact, 
								ibhact = @hact, ibuact = @uact
								where ibiddc = @tbnumi  and ibest =3
						--set @maxid1 = (select ibid from TI002 where ibiddc = @lin)
						--Detalle
						update TI0021
							set iccant =@tbcmin       --@cantE -@lccant Preguntar a Guido
								from TI0021 inner join TI002 ON TI002.ibid=TI0021.icibid AND TI002.ibiddc=@tbnumi and TI002.ibest =3
								--where icibid = @maxid1
								
					commit tran Tr_UpdateTI001
					print concat('Se actualizo el saldo del producto con codigo: ', @tbty5prod)
				end try
				begin catch
					rollback tran Tr_UpdateTI001
					print concat('No se pudo actualizo el saldo del producto con codigo: ', @tbty5prod)
				end catch
			end
			else
			begin
				begin try
					begin tran Tr_InsertTI001
					set @deposito =(Select b.abnumi   from TV001  as a,TA002 as b,TA001 as c where c.aata2depVenta =b.abnumi 
and a.tanumi  =@tbtv1numi and a.taalm  =c.aanumi )
						--Insertar Saldo Inventario
					Insert into TI001 values(@deposito ,CONVERT(int, @tbty5prod), @tbcmin, -@tbumin
					, @tblote , @tbfechavenc )
			
						--Modificar Movimiento
						--Cabecera
						Update TI002 
							set ibconcep = 	@salida, ibobs = @obs, ibfact = @fact, 
								ibhact = @hact, ibuact = @uact
								where ibiddc = @tbnumi and ibest =3 
						--set @maxid1 = (select ibid from TI002 where ibiddc = @lin)
						--Detalle
						update TI0021
							set iccant =@tbcmin      --@cantE -@lccant Preguntar a Guido
								from TI0021 inner join TI002 ON TI002.ibid=TI0021.icibid AND TI002.ibiddc=@tbnumi 
								and TI002 .ibest =3
								--where icibid = @maxid1

					commit tran Tr_InsertTI001
					print concat('Se grabo el saldo del producto con codigo: ', @tbty5prod)
				end try
				begin catch
					rollback tran Tr_InsertTI001
					print concat('No se grabo el saldo del producto con codigo: ', @tbty5prod)
				end catch
			end

			--FETCH MiCursor2 into @cpcom
	--	END
	--	CLOSE MiCursor2
--DEALLOCATE MiCursor2

	--end

	Fetch MiCursor into @tbnumi,@tbtv1numi,@tbty5prod,@tbcmin,@tbumin,@tblote,@tbfechavenc
end

--Cerrar el Cursor
close MiCursor
--Liberar la memoria
deallocate MiCursor
END
