/*==============================================================================
  Fix: dbo.SetColumnDefaultValue - thêm @@SPID vào tên constraint cho bảng tạm
  ------------------------------------------------------------------------------
  Lý do:
    Với bảng tạm (@sTableName bắt đầu bằng '#'), proc build tên constraint =
    substring(@sTableName,2,...) + '_' + column_name - KHÔNG có gì phân biệt
    session. Ví dụ nhiều session cùng gọi EXEC SetColumnDefaultValue
    '#Incd1Tmp', '' (qua dbo.INCD1_realtime_input) đồng thời sẽ cùng tạo
    constraint tên 'Incd1Tmp_ma_kho', 'Incd1Tmp_ma_vt'... Mỗi session có bảng
    tạm riêng (SQL Server tự đặt tên nội bộ khác nhau), nhưng tên CONSTRAINT
    vẫn phải duy nhất trong toàn schema tempdb.dbo -> session thứ 2 gặp lỗi
    2714 ("already an object named...") kéo theo 1750, và áp lực tranh chấp
    catalog gây thêm deadlock 1205.

    Proc này dùng chung cho 21 proc khác trong toàn hệ thống (không chỉ
    COQCS): BCKKVT, GetKetQuaKinhDoanh, INCD1/INCD1_realtime/INCD1_realtime_input/
    INCD11/INCD1A/INCD1B/Incd1lo/INCD1LOHSD, các *-CREATEPXD, sis_GetCNTT/
    sis_GetPSTK/sis_GetSDTK, ZCINCD12...

  Đã kiểm tra an toàn:
    dbo.RemoveColumnDefaultValue (proc gỡ default tương ứng) tra tên constraint
    TẠI RUNTIME qua sys.default_constraints theo parent_object_id, không đoán
    tên cố định -> không bị ảnh hưởng khi đổi cách đặt tên ở đây.
    Không có proc nào khác trong DB dự đoán/khớp cứng theo đúng pattern tên
    constraint này (đã rà soát toàn bộ sys.sql_modules).

  Thay đổi:
    CHỈ 1 chỗ - nhánh bảng tạm (if left(@sTableName,1) = '#') - thêm
    + '_' + CAST(@@SPID AS VARCHAR(10)) vào tên constraint. Nhánh bảng thật
    (permanent table) giữ nguyên, không đổi.

  Cách rollback:
    Chạy lại script "ORIGINAL (rollback)" ở cuối file này.
==============================================================================*/

ALTER PROCEDURE [dbo].[SetColumnDefaultValue]
(
	@sTableName varchar(32),
	@notUpdateColumns varchar(max)
)
as
SET TRANSACTION ISOLATION LEVEL SNAPSHOT
begin
declare @str varchar(max)

select *
into #temp
from dbo.split(@notUpdateColumns, ';')

--ALTER TABLE ph71 ADD CONSTRAINT def_ma_ct DEFAULT '' FOR ma_ct
SELECT identity(int, 1,1) as stt, 'ALTER TABLE ' + table_name + ' ADD CONSTRAINT ' + @sTableName + '_' + column_name
	+ ' DEFAULT ' + Case when data_type in('numeric', 'decimal') then '0' else '''''' end
	+ ' FOR ' + column_name as cmd
	into #cmdlist
FROM INFORMATION_SCHEMA.COLUMNS
where table_name = @sTableName and data_type in ('char', 'nchar', 'varchar', 'nvarchar', 'numeric', 'decimal')
	and column_name not in (select * from #temp)

if left(@sTableName, 1) = '#' --table temp
begin
	insert #cmdlist(cmd)
	-- [FIX SPID] thêm '_' + CAST(@@SPID AS VARCHAR(10)) để tên constraint không trùng
	-- giữa các session cùng gọi proc này trên 1 tên bảng tạm giống nhau.
	SELECT 'ALTER TABLE ' + table_name + ' ADD CONSTRAINT ' + substring(@sTableName, 2, len(@sTableName)-1) + '_' + column_name + '_' + CAST(@@SPID AS VARCHAR(10))
		+ ' DEFAULT ' + Case when data_type in('numeric', 'decimal') then '0' else '''''' end
		+ ' FOR ' + column_name
	FROM tempdb.INFORMATION_SCHEMA.COLUMNS
	where object_id('tempdb..' + table_name) = object_id('tempdb..' + @sTableName) and data_type in ('char', 'nchar', 'varchar', 'nvarchar', 'numeric', 'decimal')
		and column_name not in (select * from #temp)
end

declare @i int
set @i = 1
while @i is not null
begin
	select @str = cmd from #cmdlist where stt = @i
	print @str
	EXEC sp_sqlExec @str
	set @i = (select min(stt) from #cmdlist where stt > @i)
end

end
GO

/*==============================================================================
  ORIGINAL (rollback) - chạy khối này để khôi phục lại bản gốc nếu cần.
==============================================================================

ALTER PROCEDURE [dbo].[SetColumnDefaultValue]
(
	@sTableName varchar(32),
	@notUpdateColumns varchar(max)
)
as
SET TRANSACTION ISOLATION LEVEL SNAPSHOT
begin
declare @str varchar(max)

select *
into #temp
from dbo.split(@notUpdateColumns, ';')

SELECT identity(int, 1,1) as stt, 'ALTER TABLE ' + table_name + ' ADD CONSTRAINT ' + @sTableName + '_' + column_name
	+ ' DEFAULT ' + Case when data_type in('numeric', 'decimal') then '0' else '''''' end
	+ ' FOR ' + column_name as cmd
	into #cmdlist
FROM INFORMATION_SCHEMA.COLUMNS
where table_name = @sTableName and data_type in ('char', 'nchar', 'varchar', 'nvarchar', 'numeric', 'decimal')
	and column_name not in (select * from #temp)

if left(@sTableName, 1) = '#' --table temp
begin
	insert #cmdlist(cmd)
	SELECT 'ALTER TABLE ' + table_name + ' ADD CONSTRAINT ' + substring(@sTableName, 2, len(@sTableName)-1) + '_' + column_name
		+ ' DEFAULT ' + Case when data_type in('numeric', 'decimal') then '0' else '''''' end
		+ ' FOR ' + column_name
	FROM tempdb.INFORMATION_SCHEMA.COLUMNS
	where object_id('tempdb..' + table_name) = object_id('tempdb..' + @sTableName) and data_type in ('char', 'nchar', 'varchar', 'nvarchar', 'numeric', 'decimal')
		and column_name not in (select * from #temp)
end

declare @i int
set @i = 1
while @i is not null
begin
	select @str = cmd from #cmdlist where stt = @i
	print @str
	EXEC sp_sqlExec @str
	set @i = (select min(stt) from #cmdlist where stt > @i)
end

end
GO
==============================================================================*/
