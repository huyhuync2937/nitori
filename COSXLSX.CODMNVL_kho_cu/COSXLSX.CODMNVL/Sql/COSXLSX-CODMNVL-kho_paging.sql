/*
    Thêm phân trang kiểu "keyset" (seek theo ma_sp, ma_vt) cho proc lấy dữ liệu
    lưới "Định mức NVL theo kỳ/chuyền" (COSXLSX.CODMNVL), phục vụ scroll load từng phần.

    - Khi gọi KHÔNG truyền @PageSize (hoặc truyền 0/NULL) => trả về TOÀN BỘ dữ liệu,
      giữ nguyên hành vi cũ (tương thích ngược với các nơi khác có thể đang gọi proc này).
    - Khi truyền @PageSize > 0 => chỉ trả về tối đa @PageSize dòng kế tiếp SAU dòng có
      khóa (ma_sp, ma_vt) = (@LastMaSp, @LastMaVt). Client tự đánh lại cột "stt" theo
      thứ tự đã tải (không dùng giá trị stt do proc trả về trong nhánh phân trang).
    - @Ma_px, @Ma_vt và @Ma_sp là điều kiện lọc TÙY CHỌN phục vụ nút "Tìm kiếm" theo mã
      sản phẩm (dmsp), mã vật tư (dmvt) và mã chuyền (dmpx) trên màn hình: truyền NULL
      hoặc rỗng => không lọc theo cột tương ứng.

    Vì ma_ky/ma_px đã bị lọc cố định bằng dấu "=" (khi có truyền @Ma_px), khóa sắp xếp/
    khóa seek chỉ cần (ma_sp, ma_vt) là đủ duy nhất trong phạm vi 1 lần gọi.

    KHUYẾN NGHỊ (không tự chạy kèm ở đây, cần DBA xem xét trước khi tạo trên bảng lớn):
    tạo index hỗ trợ seek để tránh scan/sort toàn bộ dữ liệu mỗi lần gọi:

    CREATE NONCLUSTERED INDEX [IX_cosxlsx_dmdmvtct_ma_ky_px_sp_vt]
        ON [dbo].[cosxlsx-dmdmvtct] (ma_ky, ma_px, ma_sp, ma_vt)
        INCLUDE (sl_dm, ghi_chu, dvt1, stt_rec0, date0, time0, user_id0, date2, time2, user_id2);

    Nếu chưa có index này, câu SELECT (kể cả nhánh phân trang) vẫn phải join/scan qua
    view v_cosxlsx-dmdmvtct rồi mới lọc, nên tốc độ cải thiện sẽ hạn chế hơn nhiều.
*/

ALTER PROCEDURE [dbo].[COSXLSX-CODMNVL-kho]

    @Ma_ky      udt_para_ma_dm,
    @Ma_px      udt_para_ma_dm = NULL, -- NULL/rỗng = không lọc theo chuyền
    @Ma_vt      VARCHAR(20)   = NULL,  -- NULL/rỗng = không lọc theo vật tư
    @Ma_sp      VARCHAR(20)   = NULL,  -- NULL/rỗng = không lọc theo sản phẩm
    @PageSize   INT           = 0,     -- 0/NULL = trả toàn bộ (hành vi cũ)
    @LastMaSp   VARCHAR(20)   = NULL,  -- ma_sp của dòng cuối cùng đã tải (trang đầu = NULL)
    @LastMaVt   VARCHAR(20)   = NULL   -- ma_vt của dòng cuối cùng đã tải (trang đầu = NULL)

AS
SET TRANSACTION ISOLATION LEVEL SNAPSHOT
BEGIN
    SELECT '' a;

    IF ISNULL(@PageSize, 0) <= 0
    BEGIN
        -- Hành vi cũ: trả toàn bộ dữ liệu khớp ma_ky/ma_px/ma_vt/ma_sp
        SELECT ROW_NUMBER() OVER(ORDER BY a.ma_sp, a.ma_px, a.ma_vt) AS stt,
            CAST(0 AS BIT) AS chon,
            a.ma_sp+'_'+a.ma_ky+'_'+a.ma_px AS tag,
            a.ma_ky, a.stt_rec0, a.ma_sp, a.ma_vt, a.sl_dm, a.ma_px, a.ghi_chu,
            a.dvt1, a.Ten_Vt, a.Dvt, a.ten_sp, a.ten_sp2, a.ten_ky,
            a.date0, a.time0, a.user_id0, a.date2, a.time2, a.user_id2, a.nh_vt1,
            a.user_name0, a.user_name2,
            '' loai_cp
        FROM [v_cosxlsx-dmdmvtct] a
        WHERE a.ma_ky = @Ma_ky
          AND (NULLIF(LTRIM(RTRIM(@Ma_px)), '') IS NULL OR a.ma_px = @Ma_px)
          AND (NULLIF(LTRIM(RTRIM(@Ma_vt)), '') IS NULL OR a.ma_vt = @Ma_vt)
          AND (NULLIF(LTRIM(RTRIM(@Ma_sp)), '') IS NULL OR a.ma_sp = @Ma_sp)
        ORDER BY a.ma_sp, a.ma_px, a.ma_vt;
    END
    ELSE
    BEGIN
        -- Phân trang: chỉ lấy tối đa @PageSize dòng ngay sau khóa (@LastMaSp, @LastMaVt)
        SELECT TOP (@PageSize)
            0 AS stt,   -- client tự đánh lại stt tuần tự, giá trị này không dùng
            CAST(0 AS BIT) AS chon,
            a.ma_sp+'_'+a.ma_ky+'_'+a.ma_px AS tag,
            a.ma_ky, a.stt_rec0, a.ma_sp, a.ma_vt, a.sl_dm, a.ma_px, a.ghi_chu,
            a.dvt1, a.Ten_Vt, a.Dvt, a.ten_sp, a.ten_sp2, a.ten_ky,
            a.date0, a.time0, a.user_id0, a.date2, a.time2, a.user_id2, a.nh_vt1,
            a.user_name0, a.user_name2,
            '' loai_cp
        FROM [v_cosxlsx-dmdmvtct] a
        WHERE a.ma_ky = @Ma_ky
          AND (NULLIF(LTRIM(RTRIM(@Ma_px)), '') IS NULL OR a.ma_px = @Ma_px)
          AND (NULLIF(LTRIM(RTRIM(@Ma_vt)), '') IS NULL OR a.ma_vt = @Ma_vt)
          AND (NULLIF(LTRIM(RTRIM(@Ma_sp)), '') IS NULL OR a.ma_sp = @Ma_sp)
          AND ( @LastMaSp IS NULL
                OR a.ma_sp > @LastMaSp
                OR (a.ma_sp = @LastMaSp AND a.ma_vt > @LastMaVt) )
        ORDER BY a.ma_sp, a.ma_vt;
    END
END
GO
