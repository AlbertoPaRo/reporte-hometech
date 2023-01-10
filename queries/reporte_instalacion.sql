select NroDocumento  'Factura',
    WMS.NombreTitular,
    ItemCode  'Codigo SKU',
    ItemName 'SKU Identificado',
    Fecha_corta 'Fecha Venta',
    vc.CI,
    Celular,
    case
        when Sucursal = 'BR' then 'BRASIL'
        when Sucursal = 'CE' then 'EQUIPETROL'
        when Sucursal = 'VI' then 'VILLA'
        when Sucursal = 'PA' then 'PAMPA'
        when Sucursal = 'SD' then 'SANTOS DUMONT'
    end NO_Sucursal,
    WMS.FechaFactura as 'Fecha factura',
    WMS.FechaEntregaPlanner
from TB_VISTA_COMERCIAL VC
    inner join (
        select DocNum,
            CI,
            NombreTitular,
            FechaEntregaPlanner,
            FechaFactura
        from [ENT_Entrega]
    ) AS WMS ON VC.CI = WMS.CI
    and vc.Fecha = wms.FechaFactura
WHERE ItemCode IN ('SS0013', 'SS0012', 'SS0018', 'SS0019')
    and Fecha = convert(date, GETDATE() -1)
Group by NroDocumento,
    ItemCode,
    ItemName,
    Fecha_corta,
    vc.CI,
    Celular,
    Sucursal,
    WMS.FechaFactura,
    WMS.FechaEntregaPlanner,
    WMS.NombreTitular