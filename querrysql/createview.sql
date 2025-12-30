CREATE OR ALTER VIEW vw_so_kho_processed AS
WITH so_processed AS ( 
    SELECT
        [Material],
        [Sales Document],
        [Item Description],
        [Shipped Quantity (KG)] AS [Shipped Quantity (KG)],
        [Quantity (KG)] AS [Quantity (KG)],
        CASE 
            WHEN [Quantity (KG)] > 0 
                 THEN ROUND([Shipped Quantity (KG)] * 100.0 / [Quantity (KG)], 2)
            ELSE 0
        END AS Process
    FROM so
),
kho_clean AS (
    SELECT
        [Order],
        [Material],
        [SO Mapping],
        [Material Description],
        [Customer N],
        [Khối lượng]
    FROM kho
    WHERE [SO Mapping] IS NOT NULL 
      AND [SO Mapping] != '0'
)
SELECT
    k.[Order],
    k.[SO Mapping],
    k.Material,
    k.[Material Description],
    k.[Customer N],
    SUM(k.[Khối lượng]) AS [SL Mapping kho],
    s.[Shipped Quantity (KG)],
    s.[Quantity (KG)],
    s.Process
FROM kho_clean k
INNER JOIN so_processed s
    ON k.[SO Mapping] = s.[Sales Document]
   AND k.Material = s.Material
GROUP BY 
    k.[Order], k.[SO Mapping], k.Material,
    k.[Material Description], k.[Customer N],
    s.[Shipped Quantity (KG)], s.[Quantity (KG)], s.Process;




-- khách hàng:
       SELECT s.[Customer Number] AS customer_num,
               s.[Customer]          AS customer_name,
               s.[Sales Document]      AS so_mapping,
               s.[Item Description]AS material,
               s.[Reqd Deliv Date] AS req_date,
               s.[Shipped Quantity (KG)] AS shipped_qty,
               s.[Quantity (KG)]   AS qty,
               t.Process AS process_value,
               t.[SL Mapping kho] as mapping_kho,
               CASE
                 WHEN Process >= 95 THEN 'bg-success'
                 WHEN Process >= 75 THEN 'bg-warning'
                 ELSE 'bg-danger'
               END AS process_color
        FROM so s
        JOIN ton_kho t
          ON s.[Sales Document] = t.[SO Mapping]
         AND s.[Material]      = t.[Material]
        WHERE EXISTS (SELECT 1 FROM lsx x WHERE x.[Order] = t.[Order])
        ORDER BY customer_num, so_mapping;
-- của so_process ban đầu :
WITH so_processed AS (
        SELECT
            [Material],
            [Sales Document],
            [Item Description],
            [Shipped Quantity (KG)],
            [Quantity (KG)]
        FROM so
    ),
    kho_clean AS (
        SELECT
            [Order],
            [Material],
            [SO Mapping],
            [Material Description],
            [Customer N],
            [Khối lượng]
        FROM kho
        WHERE [SO Mapping] IS NOT NULL
        AND [SO Mapping] <> '0'
    )
    SELECT
        k.[Order],
        s.[Sales Document] as [SO Mapping],
        COALESCE(k.Material, s.Material)              AS Material,
        COALESCE(k.[Material Description],
                s.[Item Description])               AS [Material Description],
        k.[Customer N],
        SUM(k.[Khối lượng])                         AS [SL Mapping kho],
        s.[Shipped Quantity (KG)],
        s.[Quantity (KG)],
        /* ✅ Process cuối: (Mapping_kho + Shipped) / Quantity * 100 */
        CASE 
            WHEN s.[Quantity (KG)] > 0 
            THEN ROUND( (SUM(k.[Khối lượng]) + s.[Shipped Quantity (KG)])
                        * 100.0 / s.[Quantity (KG)], 2)
            ELSE 0
        END                                          AS Process
    FROM so_processed s
    INNER JOIN kho_clean k
            ON k.[SO Mapping] = s.[Sales Document]
            AND k.Material = s.Material
    GROUP BY 
        k.[Order],
        s.[Sales Document],
        COALESCE(k.Material, s.Material),
        COALESCE(k.[Material Description], s.[Item Description]),
        k.[Customer N],
        s.[Shipped Quantity (KG)],
        s.[Quantity (KG)];

-- khách hàng cũ
       SELECT DISTINCT
               s.[Customer Number] AS customer_num,
               s.[Customer]          AS customer_name,
               s.[Sales Document]      AS so_mapping,
               s.[Item Description]AS material,
               s.[Reqd Deliv Date] AS req_date,
               s.[Shipped Quantity (KG)] AS shipped_qty,
               s.[Quantity (KG)]   AS qty,
               t.Process AS process_value,
               t.[SL Mapping kho] as mapping_kho,
               CASE
                 WHEN Process >= 95 THEN 'bg-success'
                 WHEN Process >= 75 THEN 'bg-warning'
                 ELSE 'bg-danger'
               END AS process_color
        FROM so s
        JOIN ton_kho t
          ON s.[Sales Document] = t.[SO Mapping]
         AND s.[Material]      = t.[Material]
        ORDER BY customer_num, so_mapping;
-- ID CUỘN BÓ
SELECT distinct
            s.[ID Cuộn Bó],
            s.[Material Description],
            s.[Nhóm],
            s.[Vị trí],
            s.[Lô phôi],
            s.[Khối lượng],
            s.[Ngày sản xuất],
            CASE WHEN s.[Tp loại 2] = 1 THEN 1 ELSE 0 END AS TpLoai2,
            CASE WHEN s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL THEN N'Chờ nhập kho'
                 ELSE N'Đã nhập kho' END AS TrangThai,
            CASE WHEN s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL
                 THEN DATEADD(DAY,7,s.[Ngày sản xuất]) ELSE NULL END AS NgayDuKien
        FROM sanluong s
        WHERE s.[ID Cuộn Bó] IS NOT NULL AND s.[ID Cuộn Bó] <> ''
        AND NOT EXISTS (SELECT 1 FROM kho k WHERE k.[ID Cuộn Bó] = s.[ID Cuộn Bó])
        UNION ALL
        SELECT
            k.[ID Cuộn Bó],
            k.[Material Description],
            k.[Nhóm],
            k.[Vị trí],
            k.[Lô phôi],
            k.[Khối lượng],
            k.[Ngày sản xuất],
            CASE WHEN k.[Tp loại 2] = 'X' THEN 1 ELSE 0 END AS TpLoai2,
            CASE WHEN k.[SO Mapping] = 0 OR k.[SO Mapping] IS NULL THEN N'Nhập kho chưa mapping'
                 ELSE N'Nhập kho đã mapping' END AS TrangThai,
            CAST(NULL AS DATE) AS NgayDuKien
        FROM kho k
        WHERE k.[ID Cuộn Bó] IS NOT NULL AND k.[ID Cuộn Bó] <> ''
        --- LỊCH TÀU CŨ
        WITH tau_status AS (
    SELECT
        [TÀU/PHƯƠNG TIỆN VẬN TẢI] AS tau,
        SUM([KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU]) * 1.0 / NULLIF(MAX([KHỐI LƯỢNG TỔNG TÀU]),0) AS tau_process,
        CASE
            WHEN SUM([KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU]) * 1.0 / NULLIF(MAX([KHỐI LƯỢNG TỔNG TÀU]),0) >= 0.9 THEN 'bg-success'
            ELSE 'bg-warning'
        END AS tau_color
    FROM dbo.lichtau
    GROUP BY [TÀU/PHƯƠNG TIỆN VẬN TẢI]
),
data AS (
    SELECT
        lt.[TÀU/PHƯƠNG TIỆN VẬN TẢI] AS tau,
        s.[SO Mapping] AS saleO,
        s.[Material Description] AS material,
        lt.[KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU] AS so_khoi_luong,
        lt.[KHỐI LƯỢNG TỔNG TÀU] AS khoi_luong_tong,
        lt.[ĐẠI LÝ] AS daily,
        lt.[Cảng xếp] AS cangxep,
        lt.[CẢNG ĐẾN] AS cangden,
        ISNULL(s.[Shipped Quantity (KG)],0) AS shipped_qty,
        ISNULL(s.[Quantity (KG)],0) AS qty,
        ISNULL(s.[SL Mapping kho],0) AS Mapping_kho,
        lt.[NGÀY DK DUYỆT SO] AS duyetso,
        lt.[NHỊP] AS nhip,
        lt.SheetMonth
    FROM vw_so_kho_summary1 s
    JOIN dbo.lichtau lt
        ON s.[SO Mapping] = TRY_CAST(lt.[SỐ LỆNH TÁCH] AS BIGINT)
)
SELECT DISTINCT
    d.*,
    ROUND( (Mapping_kho + shipped_qty) * 100.0 / NULLIF(qty,0), 2 ) AS process_value,
    CASE
        WHEN (Mapping_kho + shipped_qty) * 100.0 / NULLIF(qty,0) >= 90 THEN 'bg-success'
        WHEN (Mapping_kho + shipped_qty) * 100.0 / NULLIF(qty,0) >= 75 THEN 'bg-warning'
        ELSE 'bg-danger'
    END AS process_color,
    ts.tau_process,
    ts.tau_color AS tau_status_color
FROM data d
LEFT JOIN tau_status ts
    ON d.tau = ts.tau