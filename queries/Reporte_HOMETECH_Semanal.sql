select NroDocumento 'Numero de Documento',
    ItemCode 'Item',
    Semana,
    ItemName 'Descripcion',
    CONVERT(INT, Cantidad) Cantidad,
    replace((Monto_real) * 0.87, '.', ',') 'Precio S/IVA',
    (dateadd(hour, 23, Fecha)) 'Fecha',
    Nombre_Vendedor 'Vendedor',
    case
        when Sucursal = 'BR' then 'BRASIL'
        when Sucursal = 'CE' then 'EQUIPETROL'
        when Sucursal = 'VI' then 'VILLA'
        when Sucursal = 'PA' then 'PAMPA'
        when Sucursal = 'SD' then 'SANTOS DUMONT'
        when Sucursal = 'TO' then 'TIENDA ONLINE'
    end Sucursal
from TB_VISTA_COMERCIAL
where Familia = 'coccion'
    and Fabricante = 'HOMETECH'
    and fecha >= convert(date,'10-06-2022')
    and fecha <= CONVERT(date, GETDATE()-2)
order by Fecha,
    Sucursal