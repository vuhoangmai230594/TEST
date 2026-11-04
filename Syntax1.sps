* Encoding: UTF-8.
* =========================================================================.
* PHẦN 1: CHUẨN HÓA MÃ PHÂN TỔ & CHỈ TIÊU TỪ FILE MICRODATA LFS
* =========================================================================.
WEIGHT BY Final_Weight_11.
*Tao biến HĐKT.
IF  ((C19= 1| C27=4| C28 = 1 | C28= 2|C29=2| C30=1)) HDKT19=1.
IF  ((C31 = 1 | C32= 1) & C36=1) HDKT19=2.
EXECUTE.
DO IF (C04 >= 15).
RECODE HDKT19 (MISSING=3).
END IF.
IF  (c04 >=15 & c06=2 ) HDKT19=4.
EXECUTE.
IF (HDKT19 <= 3 & C04 >= 15) DS15 = 1.
RECODE HDKT19 (1 thru 2 = 1) INTO LLLD19.
RECODE HDKT19 (1 = 1)        INTO CoVL19.
RECODE HDKT19 (2 = 1)        INTO Thatnghiep19.
VALUE LABELS HDKT19 
    1 "Có việc làm"
    2 "Thất nghiệp"
    3 "Không hoạt động kinh tế"
    4 "Đang sống ở nước ngoài (Abroad)".
VALUE LABELS Thatnghiep19
     1 "Thất nghiệp".
VALUE LABELS CoVL19
     1 "Có việc làm".
VALUE LABELS DS15
    1 "Dân số từ 15 tuổi trở lên".
VARIABLE LABELS
    HDKT19             "Tình trạng hoạt động kinh tế (ICLS-19)"
    DS15               "Dân số từ 15 tuổi trở lên (Trừ người đang sống ở nước ngoài)"
    LLLD19             "Lực lượng lao động (ICLS-19)"
    CoVL19             "Lao động có việc làm (ICLS-19)"
    Thatnghiep19       "Lao động thất nghiệp (ICLS-19)".
EXECUTE.
* -------------------------------------------------------------------------.
* KHỐI 1: THỜI GIAN (year, quarter)
* -------------------------------------------------------------------------.
COMPUTE year = NUMBER(NamDT, F4.0).
COMPUTE quarter = TRUNC((ThangDT - 1) / 3) + 1.
FORMATS year quarter (F4.0).
EXECUTE.
* -------------------------------------------------------------------------.
* KHỐI 2: ĐỊA BÀN (vn_code, vung_code, tinh_code) TỪ BIẾN MaTinh
* -------------------------------------------------------------------------.
STRING vn_code vung_code tinh_code (A2).
COMPUTE vn_code = 'VN'.
COMPUTE tinh_code = STRING(MaTinh, N2).

* Ánh xạ MaTinh sang 6 Vùng KT-XH (chuẩn cho cả hệ 63 tỉnh và 34 tỉnh).
IF ANY(MaTinh, 2, 4, 6, 8, 10, 11, 12, 14, 15, 17, 19, 20, 24, 25) vung_code = 'V1'.
IF ANY(MaTinh, 1, 22, 26, 27, 30, 31, 33, 34, 35, 36, 37)          vung_code = 'V2'.
IF ANY(MaTinh, 38, 40, 42, 44, 45, 46, 48, 49, 51, 52, 54, 56, 58, 60) vung_code = 'V3'.
IF ANY(MaTinh, 62, 64, 66, 67, 68)                                 vung_code = 'V4'.
IF ANY(MaTinh, 70, 72, 74, 75, 77, 79)                             vung_code = 'V5'.
IF ANY(MaTinh, 80, 82, 83, 84, 86, 87, 89, 91, 92, 93, 94, 95, 96) vung_code = 'V6'.
EXECUTE.
* -------------------------------------------------------------------------.
* KHỐI 3: CÁC BIẾN PHÂN TỔ CHUẨN TEMPLATE EXCEL (LẤY TỪ C02, C04, TTNT...)
* -------------------------------------------------------------------------.
STRING area (A5) sex (A3) age_group (A5) age_group_36 (A5) 
       sector_str (A3) job_status (A8) informal_status (A8).

* 1. Khu vực (TTNT: 1=Thành thị, 2=Nông thôn).
IF (TTNT = 1) area = 'Urban'.
IF (TTNT = 2) area = 'Rural'.

* 2. Giới tính (Lấy từ C02: 1=Nam, 2=Nữ).
IF (C02 = 1) sex = 'M'.
IF (C02 = 2) sex = 'F'.

* 3. Nhóm tuổi (Lấy từ C04 do biến AGE và NHOMTUOI7 đã bỏ).
IF (C04 >= 15 AND C04 <= 24) age_group = '15-24'.
IF (C04 >= 25 AND C04 <= 34) age_group = '25-34'.
IF (C04 >= 35 AND C04 <= 44) age_group = '35-44'.
IF (C04 >= 45 AND C04 <= 54) age_group = '45-54'.
IF (C04 >= 55 AND C04 <= 64) age_group = '55-64'.
IF (C04 >= 65)               age_group = '65+'.

* Nhóm tuổi riêng cho Biểu 36 (chỉ từ 15-24 đến 55-64).
IF (C04 >= 15 AND C04 <= 64) age_group_36 = age_group.

* 4. Nhóm ngành kinh tế (nganhkt: 1=NLTS, 2=CN-XD, 3=Dịch vụ).
IF (nganhkt = 1) sector_str = 'AGR'.
IF (nganhkt = 2) sector_str = 'IND'.
IF (nganhkt = 3) sector_str = 'SER'.

* 5. Vị thế việc làm (C49: 5=Làm công, 2=Tự làm, 1=Chủ cơ sở, 4=LĐ gia đình).
IF (C49 = 5) job_status = 'EMPLOYEE'.
IF (C49 = 2) job_status = 'SELF'.
IF (C49 = 1) job_status = 'EMPLOYER'.
IF (C49 = 4) job_status = 'FAMILY'.

* 6. Chính thức / Phi chính thức (informal_agr: 1=Chính thức, 2=Phi chính thức).
IF (informal_agr = 1) informal_status = 'FORMAL'.
IF (informal_agr = 2) informal_status = 'INFORMAL'.
EXECUTE.

* -------------------------------------------------------------------------.
* KHỐI 4: TẠO LẠI 5 BIẾN CHỈ TIÊU 
* -------------------------------------------------------------------------.

* * Thiếu việc làm theo thời gian thực tế (Làm dưới 35 giờ/tuần, muốn và sẵn sàng làm thêm).
NUMERIC Thieuvieclam (F1.0).
IF ((C65 >= 0 & C65 < 35) & C67 = 1 & C68 = 1 & HDKT19 = 1) Thieuvieclam = 1.
DO IF (HDKT19 = 1).
    RECODE Thieuvieclam (MISSING = 9).
END IF.

* 2. Thanh niên NEET 15-24 tuổi (Biểu 18):
*    Tuổi từ 15-24 + Không có việc làm (HDKT19 = 2 hoặc 3) + Không đi học/đào tạo (C19 = 2).
NUMERIC neet (F1.0).
IF (C04 >= 15 & C04 <= 24 & (C13 = 2 & C14 = 2) & (HDKT19 = 2 | HDKT19 = 3)) neet = 1.
DO IF (C04 >= 15 & C04 <= 24 & (HDKT19 = 1 | HDKT19 = 2 | HDKT19 = 3)).
    RECODE neet (MISSING = 2).
END IF.

* 3. Lực lượng lao động tiềm năng (Biểu 33 - Potential Labour Force):
*    Ngoài LLLĐ (HDKT19=3) nhưng: (Tìm việc C31=1 song chưa sẵn sàng C34=2) HOẶC (Không tìm việc C31=2 nhưng muốn làm C32=1 và sẵn sàng C34=1).
COMPUTE LLLD_tiemnang = $SYSMIS.
IF (HDKT19 = 3 AND ((C31 = 1 AND C34 = 2) OR (C31 = 2 AND C32 = 1 AND C34 = 1))) LLLD_tiemnang = 1.

* 4. Lực lượng lao động mở rộng (Extended Labour Force = LLLD19 + LLLD_tiemnang).
COMPUTE LLLD_morong = $SYSMIS.
IF (LLLD19 = 1 OR LLLD_tiemnang = 1) LLLD_morong = 1.

* * Tạo biến lao động tiềm năng và không sử dụng hết tiềm năng.
NUMERIC tiemnang (F1.0) khongsdhet (F1.0).
IF (HDKT19 = 3 & (((C31 = 1 | C32 = 1) & C36 = 2) | (C31 = 2 & C36 = 1))) tiemnang = 1.
IF (HDKT19 = 2 | Thieuvieclam = 1 | tiemnang = 1) khongsdhet = 1.
EXECUTE.

* -------------------------------------------------------------------------.
* KHỐI 5: CÁC BIẾN ĐỊNH LƯỢNG (SCALE) ĐỂ TÍNH MEAN / MEDIAN TRONG CTABLES
* -------------------------------------------------------------------------.

* Thu nhập tháng (đổi từ nghìn đồng sang đồng trên người có việc làm có thu nhập C66 > 0).
COMPUTE luong_thang = $SYSMIS.
IF (CoVL19 = 1 AND C66 > 0) luong_thang = C66 * 1000.

* Các biến cờ 0/100 để tính Tỷ lệ (%) bằng hàm [MEAN]:
* Biểu 03: Tỷ lệ tham gia LLLĐ (LLLD19 = 1 trên DS15 = 1).
COMPUTE ind_03 = $SYSMIS.
IF (DS15 = 1) ind_03 = 0.
IF (DS15 = 1 AND LLLD19 = 1) ind_03 = 100.

* Biểu 04: Tỷ lệ thất nghiệp (Thatnghiep19 = 1 trên LLLD19 = 1).
COMPUTE ind_04 = $SYSMIS.
IF (LLLD19 = 1) ind_04 = 0.
IF (LLLD19 = 1 AND Thatnghiep19 = 1) ind_04 = 100.

* Biểu 09: Tỷ lệ có việc làm / Dân số 15+ (CoVL19 = 1 trên DS15 = 1).
COMPUTE ind_09 = $SYSMIS.
IF (DS15 = 1) ind_09 = 0.
IF (DS15 = 1 AND CoVL19 = 1) ind_09 = 100.

* Biểu 10: Tỷ lệ làm công hưởng lương (C49 = 5 trên CoVL19 = 1).
COMPUTE ind_10 = $SYSMIS.
IF (CoVL19 = 1) ind_10 = 0.
IF (CoVL19 = 1 AND C49 = 5) ind_10 = 100.

* Biểu 11: Tỷ lệ việc làm NLTS (nganhkt = 1 trên CoVL19 = 1).
COMPUTE ind_11 = $SYSMIS.
IF (CoVL19 = 1) ind_11 = 0.
IF (CoVL19 = 1 AND nganhkt = 1) ind_11 = 100.

* Biểu 15: Tỷ lệ thiếu việc làm chung (Thieuvieclam = 1 trên CoVL19 = 1).
COMPUTE ind_15 = $SYSMIS.
IF (CoVL19 = 1) ind_15 = 0.
IF (CoVL19 = 1 AND Thieuvieclam = 1) ind_15 = 100.

* Biểu 16: Tỷ lệ thất nghiệp thanh niên 15-24.
COMPUTE ind_16 = $SYSMIS.
IF (LLLD19 = 1 AND C04 >= 15 AND C04 <= 24) ind_16 = 0.
IF (Thatnghiep19 = 1 AND C04 >= 15 AND C04 <= 24) ind_16 = 100.

* Biểu 17: Tỷ lệ thất nghiệp trong độ tuổi lao động (tuoild = 1).
COMPUTE ind_17 = $SYSMIS.
IF (LLLD19 = 1 AND tuoild = 1) ind_17 = 0.
IF (Thatnghiep19 = 1 AND tuoild = 1) ind_17 = 100.

* Biểu 18: Tỷ lệ NEET 15-24.
COMPUTE ind_18 = $SYSMIS.
IF (DS15 = 1 AND C04 >= 15 AND C04 <= 24) ind_18 = 0.
IF (neet = 1) ind_18 = 100.

* Biểu 25: Tỷ lệ việc làm phi chính thức (informal_agr = 2).
COMPUTE ind_25 = $SYSMIS.
IF (CoVL19 = 1 AND ANY(informal_agr, 1, 2)) ind_25 = 0.
IF (CoVL19 = 1 AND informal_agr = 2) ind_25 = 100.

* Biểu 26: Tỷ lệ bao phủ phân loại phi chính thức.
COMPUTE ind_26 = $SYSMIS.
IF (CoVL19 = 1) ind_26 = 0.
IF (CoVL19 = 1 AND ANY(informal_agr, 1, 2)) ind_26 = 100.

* Biểu 32: Tỷ lệ lao động không sử dụng hết tiềm năng (Khong_sd_het = 1 trên LLLD_morong = 1).
COMPUTE ind_32 = $SYSMIS.
IF (LLLD_morong = 1) ind_32 = 0.
IF (LLLD_morong = 1 AND khongsdhet = 1) ind_32 = 100.

* Biểu 34: Tỷ lệ lao động qua đào tạo có bằng, chứng chỉ (quadaotao = 1 trên LLLD19 = 1).
COMPUTE ind_34 = $SYSMIS.
IF (LLLD19 = 1) ind_34 = 0.
IF (LLLD19 = 1 AND quadaotao = 1) ind_34 = 100.

* Biểu 36: Tỷ lệ thiếu việc làm trong độ tuổi lao động (tuoild = 1).
COMPUTE ind_36 = $SYSMIS.
IF (CoVL19 = 1 AND tuoild = 1) ind_36 = 0.
IF (CoVL19 = 1 AND tuoild = 1 AND Thieuvieclam = 1) ind_36 = 100.

* Khai báo thang đo chuẩn cho CTABLES.
VARIABLE LEVEL year quarter vn_code vung_code tinh_code 
               area sex age_group age_group_36 sector_str job_status informal_status (NOMINAL).
VARIABLE LEVEL luong_thang ind_03 ind_04 ind_09 ind_10 ind_11 ind_15 
               ind_16 ind_17 ind_18 ind_25 ind_26 ind_32 ind_34 ind_36 (SCALE).
EXECUTE.

* =========================================================================.
* PHẦN 2: CHẠY CTABLES, BẮT KẾT QUẢ BẰNG OMS VÀ XUẤT EXCEL 36 BIỂU
* =========================================================================.
* Thiết lập để CTABLES hiển thị mã (Code) thay vì Nhãn (Label) để Excel đọc được.
SET TVARS=NAMES TNUMBERS=VALUES.
DATASET NAME MicroLFS.
* Đường dẫn file.
FILE HANDLE OutExcel /NAME='C:\Users\vmhoang\Desktop\Dashboard\3_Output.xlsx'.

* -------------------------------------------------------------------------.
* BIỂU 01: LỰC LƯỢNG LAO ĐỘNG (Không phân tổ)
* -------------------------------------------------------------------------.

DATASET ACTIVATE MicroLFS.
DATASET DECLARE B01.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B01' /TAG='t01'.
TEMPORARY.
SELECT IF (LLLD19 = 1).
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) [COUNT F40.0]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE.
OMSEND TAG=['t01'].
* Xử lý Dataset kết quả và xuất đè vào sheet 01_LLLD.
DATASET ACTIVATE B01.
RENAME VARIABLES (Var1 Var2 Var3 Var4 = year quarter geo_code value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='01_LLLD'
  /KEEP = year quarter geo_code value.
DATASET CLOSE B01.

* -------------------------------------------------------------------------.
* BIỂU 02: NGƯỜI CÓ VIỆC LÀM (Phân tổ 1 chiều: Khu vực)
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B02.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B02' /TAG='t02'.
TEMPORARY.
SELECT IF (CoVL19 = 1).
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code area DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) > area [COUNT F40.0]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=area ['Urban', 'Rural'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE.
OMSEND TAG=['t02'].

DATASET ACTIVATE B02.
RENAME VARIABLES (Var1 Var2 Var3 Var4 Var5 = year quarter geo_code area value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='02_CO_VIEC_LAM'
  /KEEP = year quarter geo_code area value.
DATASET CLOSE B02.

* -------------------------------------------------------------------------.
* BIỂU 03: TỶ LỆ THAM GIA LỰC LƯỢNG LAO ĐỘNG (Phân tổ 1 chiều: Giới tính)
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B03.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B03' /TAG='t03'.
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code sex ind_03 DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) > sex BY ind_03 [MEAN F40.6]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=sex ['M', 'F'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE.
OMSEND TAG=['t03'].

DATASET ACTIVATE B03.
RENAME VARIABLES (Var1 Var2 Var3 Var4 V8 = year quarter geo_code sex value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='03_TL_THAM_GIA'
  /KEEP = year quarter geo_code sex value.
DATASET CLOSE B03.

* -------------------------------------------------------------------------.
* BIỂU 05: THU NHẬP BÌNH QUÂN (Phân tổ 4 chiều: Giới tính x Khu vực x Ngành x Phi chính thức)
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B05.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B05' /TAG='t05'.
TEMPORARY.
SELECT IF (luong_thang > 0).
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code sex area sector_str informal_status luong_thang DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) > sex > area > sector_str > informal_status BY luong_thang [MEAN F40.0]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=sex ['M', 'F'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=area ['Urban', 'Rural'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=sector_str ['AGR', 'IND', 'SER'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=informal_status ['FORMAL', 'INFORMAL'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE.
OMSEND TAG=['t05'].

DATASET ACTIVATE B05.
RENAME VARIABLES (Var1 Var2 Var3 Var4 Var5 Var6 Var7 V8 = year quarter geo_code sex area sector informal_status value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='05_LUONG_BINH_QUAN'
  /KEEP = year quarter geo_code sex area sector informal_status value.
DATASET CLOSE B05.

* -------------------------------------------------------------------------.
* BIỂU 06: DÂN SỐ TỪ 15 TUỔI TRỞ LÊN (Không phân tổ)
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B06.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B06' /TAG='t06'.
TEMPORARY.
SELECT IF (DS15 = 1).
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) [COUNT F40.0]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE.
OMSEND TAG=['t06'].

DATASET ACTIVATE B06.
RENAME VARIABLES (Var1 Var2 Var3 Var4 = year quarter geo_code value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='06_DAN_SO_15_PLUS'
  /KEEP = year quarter geo_code value.
DATASET CLOSE B06.











* -------------------------------------------------------------------------.
* CẤU TRÚC 2: CÁC BIỂU PHÂN TỔ 1 CHIỀU (VD: Biểu 02 - Khu vực, Biểu 03 - Giới tính)
* Có cột: year, quarter, geo_code, [biến phân tổ], value.
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B03.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B03' /TAG='t03'.

* Dùng biến tỷ lệ ind_03 (0/100) đã tạo ở Phần 1 để tính tự động bằng [MEAN].
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code sex ind_03 DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) > sex BY ind_03 [MEAN F40.6]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=sex ['M', 'F'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE.

OMSEND TAG=['t03'].

DATASET ACTIVATE B03.
RENAME VARIABLES (Var1 Var2 Var3 Var4 Mean = year quarter geo_code sex value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='03_TL_THAM_GIA'
  /KEEP = year quarter geo_code sex value.
DATASET CLOSE B03.


* -------------------------------------------------------------------------.
* CẤU TRÚC 3: CÁC BIỂU PHÂN TỔ 2 CHIỀU (VD: Biểu 04 - Tỷ lệ thất nghiệp theo Giới tính x Tuổi)
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B04.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B04' /TAG='t04'.

CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code sex age_group ind_04 DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) > sex > age_group BY ind_04 [MEAN F40.6]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=sex ['M', 'F'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=age_group ['15-24', '25-34', '35-44', '45-54', '55-64', '65+'] 
              EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE.

OMSEND TAG=['t04'].

DATASET ACTIVATE B04.
RENAME VARIABLES (Var1 Var2 Var3 Var4 Var5 Mean = year quarter geo_code sex age_group value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='04_TL_THAT_NGHIEP'
  /KEEP = year quarter geo_code sex age_group value.
DATASET CLOSE B04.


* -------------------------------------------------------------------------.
* CẤU TRÚC 4: BIỂU ĐA CHIỀU PHỨC TẠP (VD: Biểu 05, 25 - 4 chiều phân tổ)
* Tính Thu nhập bình quân: Giới tính x Khu vực x Ngành x Tình trạng Phi chính thức
* -------------------------------------------------------------------------.
DATASET ACTIVATE MicroLFS.
DATASET DECLARE B05.
OMS /SELECT TABLES /IF COMMANDS=['CTABLES'] SUBTYPES=['Custom Table'] 
    /DESTINATION FORMAT=SAV OUTFILE='B05' /TAG='t05'.

* Lọc những người có thu nhập > 0 trước khi tính Mean.
TEMPORARY.
SELECT IF (luong_thang > 0).
CTABLES
  /VLABELS VARIABLES=year quarter vn_code vung_code tinh_code sex area sector_str informal_status luong_thang DISPLAY=NONE
  /TABLE year > quarter > (vn_code + vung_code + tinh_code) > sex > area > sector_str > informal_status BY luong_thang [MEAN F40.0]
  /SLABELS VISIBLE=NO
  /CATEGORIES VARIABLES=year quarter vn_code vung_code tinh_code ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=sex ['M', 'F'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=area ['Urban', 'Rural'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=sector_str ['AGR', 'IND', 'SER'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE
  /CATEGORIES VARIABLES=informal_status ['FORMAL', 'INFORMAL'] EMPTY=INCLUDE TOTAL=YES LABEL='ALL' POSITION=BEFORE.

OMSEND TAG=['t05'].

DATASET ACTIVATE B05.
RENAME VARIABLES (Var1 Var2 Var3 Var4 Var5 Var6 Var7 Mean = year quarter geo_code sex area sector informal_status value).
SELECT IF NOT MISSING(value).
ALTER TYPE year quarter (F4.0).
SAVE TRANSLATE OUTFILE=OutExcel
  /TYPE=XLS /VERSION=12 /FIELDNAMES VALUE=NAMES /CELLS=VALUES /REPLACE
  /EXCELOPTIONS SHEET='05_LUONG_BINH_QUAN'
  /KEEP = year quarter geo_code sex area sector informal_status value.
DATASET CLOSE B05.




