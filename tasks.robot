*** Settings ***
Library    RPA.Browser.Selenium
Library    Collections
Library    RPA.PDF
Library    RPA.FileSystem
Library    RPA.Desktop
Library    String
Library    RPA.Excel.Application
Library    RPA.Excel.Files
Library    RPA.Smartsheet
Library    RPA.Tables
Library    XML
Library    OperatingSystem
Library    RPA.Email.ImapSmtp

*** Variable ***
${FILENAME}=    Invoice

*** Tasks ***
Open the robot order website
    Set Download Directory    /HuongHNU/HOA DON/Download    // cái ni nếu mình clone về thì máy mình có chạy được không ?
    Open Available Browser    url=https://www.meinvoice.vn/tra-cuu/
Enter code
    Input Text    //input[@id="txtCode"]    3PTWUJJ3XGB
    Click Element    //input[@id="btnSearchInvoice"]
    
     

Download Invoice  
    Set Download Directory    /HuongHNU/HOA DON/Download    // vì răng trên kia set rồi mà răng dưới ni lại lòi ra dòng ni nữa ?
    Click Element    //div[@class="res-btn download"]
    Wait Until Element Is Visible    //div[@class="dm-item xml txt-download-xml"]
    Click Element    //div[@class="dm-item xml txt-download-xml"]    

Wait For Download To Complete
    Wait Until Keyword Succeeds    5 sec    10 sec    ${FILENAME}    // bạn có chắc là hàm ni hắn chạy được không ?


Read XML 
    ${XML}=    Open File    /HuongHNU/HOA DON/Download/1C23TTT_00002034.xml    // bạn tải nhiều hóa đơn mà bạn gán cứng tên file rồi đến lúc tải nhiều mã khác nhau thì tên file đó có tồn tại để xử lý không ?
    ${xml}=    Parse Xml    /HuongHNU/HOA DON/Download/1C23TTT_00002034.xml
    ${Ma_code}=    Get Elements Texts    ${XML}    DLHDon/Id
    ${KHHDon}=    Get Elements Texts    ${XML}    DLHDon/TTChung/KHHDon
    ${SHDon}=    Get Elements Texts    ${XML}    DLHDon/TTChung/SHDon
    ${Nlap}=    Get Elements Texts    ${XML}    DLHDon/TTChung/NLap
    ${MST_invoice}=    Get Elements Texts    ${XML}    DLHDon/TTChung/MSTTCGP 
    ${MST_nguoi_ban}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/NBan/MST
    ${Ten_nguoi_ban}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/NBan/Ten
    ${Dchi_nguoi_ban}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/NBan/DChi 
    ${Ten_nguoi_mua}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/NMua/Ten 
    ${MST_nguoi_mua}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/NMua/MST
    ${Dchi_nguoi_mua}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/NMua/DChi   
    ${Ma_HHDVu}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/DSHHDVu/HHDVu/MHHDVu
    ${Ten_HHDVu}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/DSHHDVu/HHDVu/THHDVu
    ${DVTinh_HHDVu}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/DSHHDVu/HHDVu/DVTinh   
    ${SLuong_HHDVu}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/DSHHDVu/HHDVu/SLuong  
    ${DGia_HHDVu}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/DSHHDVu/HHDVu/DGia  
    ${ThTien_HHDVu}=    Get Elements Texts    ${XML}    DLHDon/NDHDon/DSHHDVu/HHDVu/ThTien
    Create Workbook    /HuongHNU/invoice.xlsx
    Save Workbook
    Open File    /HuongHNU/invoice.xlsx
    &{Employees_Row1}=    Create Dictionary    KHHDon=${KHHDon}[-1]    SHDon=${SHDon}[-1]    NLap=${Nlap}[-1]    
    ...    MSTInvoice=${MST_invoice}[-1]    MSTNguoiban=${MST_nguoi_ban}[-1]    TenNguoiBan=${Ten_nguoi_ban}[-1]    
    ...    DiachiNguoiBan=${Dchi_nguoi_ban}[-1]    TenNguoiMua=${Ten_nguoi_mua}[-1]    MSTNguoiMua=${MST_nguoi_mua}[-1]    
    ...    DChiNguoiMua=${Dchi_nguoi_mua}[-1]    MaHHDVu=${Ma_HHDVu}[-1]    TenHHDVu=${Ten_HHDVu}[-1]    
    ...    DViTinh=${DVTinh_HHDVu}[-1]    SLuongTinh=${SLuong_HHDVu}[-1]    DGia=${DGia_HHDVu}[-1]    TTien=${ThTien_HHDVu}[-1]        
    # &{Employees_Row2}=    Create Dictionary    name=John    age=${22}
    # &{Employees_Row3}=    Create Dictionary    name=Adam    age=${67}
    @{Worksheet_Data}=    Create List    ${Employees_Row1}    
    Create Worksheet
    ...    name=Invoice
    ...    content=${Worksheet_Data}
    ...    header=True
    Save Workbook
    
