// Dew.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#import "C:/Program Files/Common Files/System/ado/msado15.dll" rename("EOF", "adoEOF") rename("BOF", "adoBOF") no_namespace
//#import "C:\Program Files\Common Files\System\ado\msadox.dll" no_namespace

//#include <afxdb.h>
//#include <afxdi>
#include<stdio.h>
#include<icrsint.h> 
#include<conio.h>
#include<iostream>
#include<atlstr.h>
#include<ctime>
#include<chrono>
#include<string>
#include<fstream>
#include<comutil.h>
#include<limits>
#include<clocale>
#include<vector>
#include<opencv2/opencv.hpp>
#include<opencv2/imgproc/imgproc.hpp>
//#include<Magick++.h>
bool userVrfy(CString sdata[][100],int imax,CString T);
void viewData(CString sData[][100],int imax);
void login(CString sData[][100], int imax, _RecordsetPtr pRecordset, _ConnectionPtr pConnection);
void work(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection);
void logout();
CString gendate();
void searchDB(_RecordsetPtr   pRecordset, _ConnectionPtr  pConnection);
void deleteRec(_RecordsetPtr   pRecordset, _ConnectionPtr  pConnection);
unsigned char* magicFunc();
void magicDisplay(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection);
//void insImg(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection, _StreamPtr pStream);
//std::vector<char> ReadBMP();
CString sConn;
cv::Mat image;
int logflag = 0;
HRESULT hRet = NULL;
int m_width=0;
int m_height=0;
/*struct Color {
    float r, g, b;
    Color();
    ~Color();
};*/
//std::vector<Color> fetchOLE(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection);
//void exportBMP(const char* path, std::vector<Color> m_colors);
//Color::~Color() = default;
//Color::Color() = default;
int main()
{    //setlocale(LC_ALL, "en_US.UTF-8");
    _ConnectionPtr  pConnection=NULL;
    _CommandPtr     pCommand;
    _ParameterPtr   pParameter;
    _RecordsetPtr   pRecordset=NULL;
    _StreamPtr      pStream = NULL;
    int iErrorCode;
    HRESULT hr;
    int choose = 0;
    int imax = 0;
    CString sData[10][100];
    sConn.Format(L"Provider=Microsoft.JET.OLEDB.4.0;Data Source=%s;", L"Database11.mdb");
    //      Initialize COM  
    if (FAILED(hr = CoInitialize(NULL)))
    {
        goto done_err;
    }

    //      Intialize the ADO Connection object
    if (FAILED(hr = pConnection.CreateInstance(__uuidof(Connection))))
    {
        goto done_err;
    }

    //      Intialize the ADO Command object
    if (FAILED(hr = pCommand.CreateInstance(__uuidof(Command))))
    {
        goto done_err;
    }

    //      Intialize the ADO Parameter object
    if (FAILED(hr = pParameter.CreateInstance(__uuidof(Parameter))))
    {
        goto done_err;
    }

    //      Intialize the ADO RecordSet object
    if (FAILED(hr = pRecordset.CreateInstance(__uuidof(Recordset))))
    {
        goto done_err;
    }
    if (FAILED(hr = pStream.CreateInstance(__uuidof(Stream))))
    {
        goto done_err;
    }
    pConnection->Open(_bstr_t(sConn), _bstr_t(""), _bstr_t(""), 0);
    while (1) {
        pRecordset = pConnection->Execute("SELECT Count(*) AS Expr1 FROM Dew1; ", NULL, 1);
        imax = _ttoi((LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Expr1")->Value);
        pRecordset = pConnection->Execute("SELECT * FROM Dew1;", NULL, 1);
        pRecordset->MoveFirst();
        for (int i = 0; i < imax; i++)
        {
            sData[0][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"ID")->Value;
            sData[1][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Name")->Value;
            sData[2][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Image_ID")->Value;
            sData[3][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Date_Accessed")->Value;
            sData[4][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Operations")->Value;
            sData[5][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Image")->Value;
            //std::wcout <<(LPCTSTR)sData[0][i]<<"|"<< (LPCTSTR)sData[1][i]<<"|"<< (LPCTSTR)sData[2][i]<<"|"<< (LPCTSTR)sData[3][i] <<"|" << (LPCTSTR)sData[4][i] << "\n";
            pRecordset->MoveNext();
        }
        if (logflag == 0) {
            login(sData, imax, pRecordset, pConnection);
        }
        //imax++;
        else if (logflag == 1) {
            std::cout << "\n0:Logout\n1:Work\n2:Search\n3:delete\n4:ReadImage\n5:export";
            std::cin >> choose;
            if (choose == 0) {
                logout();
                break;
            }
            else if (choose == 1) {
                work(sData, imax, pRecordset, pConnection);
            }
            else if (choose == 2) {
                searchDB(pRecordset, pConnection);
            }
            else if (choose == 3) {
                deleteRec(pRecordset, pConnection);
            }
            else if (choose == 4)
            {
                magicFunc();
            }
            else if (choose == 5)
            {
                magicDisplay(sData, imax, pRecordset, pConnection);
            }
        }
        //viewData(sData, imax);
    }
    CoUninitialize();
    iErrorCode = 0;
done:
    std::cout <<"\n" << iErrorCode;
    return iErrorCode;
done_err:
    // TODO: Cleanup
    iErrorCode = (int)hr;
    goto done;
}
bool userVrfy(CString sData[][100],int imax,CString T) {
    
    for (int i = 0; i < imax; i++)
    {
        if (sData[1][i] == T) {
            return true;
        }
    }
    return false;

}
void viewData(CString sData[][100],int imax) {
    for (int i = 0; i < imax; i++)
    {
        std::wcout << (LPCTSTR)sData[0][i] << "|" << (LPCTSTR)sData[1][i] << "|" << (LPCTSTR)sData[2][i] << "|" << (LPCTSTR)sData[3][i] << "|" << (LPCTSTR)sData[4][i] << "\n";
    }
}
void login(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection)
{
    logflag = 1;
    CString T,img;
    char tc[1024],imgc[1024];
    std::cout << "Enter Username";
    std::cin >> tc;
    T = CString(tc);
    std::wcout << (LPCTSTR)T << "\n";
    std::cout << "Enter Image ID";
    std::cin >> imgc;
    img = CString(imgc);
    //std::wcout << (LPCTSTR)img << "\n";
    CString cimax;
    cimax.Format(L"%d", _ttoi(sData[0][imax-1]) + 1);
    unsigned char* magic = magicFunc();
    //std::cout << "\n MAGIC\n" << magic<<"\n";
    CString imgdata(magic);
    imgdata.Replace(_T("'"), _T("''"));
    /*for (int i = 0; i < data.size(); i++)
    {   if(data[i]=='\0'){
        imgdata = imgdata + '0';
    }
        else{
        CString peekdata = data[i];
        imgdata = imgdata + data[i];}
    }*/
    //std::wcout << (_bstr_t)(LPCTSTR)imgdata<<"\nIMGDATA";
    if (userVrfy(sData, imax,T) == 1)
    {
        CString str = "INSERT INTO Dew1 VALUES(" + cimax + ",'"+T+"','"+img+"','"+gendate()+"', 'l','" + imgdata + "');";
        //std::wcout << (_bstr_t)(LPCTSTR)str << "\n";
        //CString cstr = CString(str.c_str());
        std::cout << "User Verified.\n";
        pConnection->Execute(_bstr_t(LPCTSTR(str)), NULL, 1);
    }
    else
    {
        std::cout << "Error.Please Enter Valid Username.\n";

    }
}
void work(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection) {
    CString op,cimax;
    char opc[1024];
    std::cout << "Enter Operation" << "\n";
    std::cin >> opc;
    op = CString(opc);
    cimax.Format(L"%d", _ttoi(sData[0][imax-1]));
    CString str = "UPDATE Dew1 SET Operations = (Operations + ',"+op+"') WHERE ID = "+cimax+";";
    //std::wcout << (LPCTSTR)str;
    pConnection->Execute(_bstr_t(LPCTSTR(str)), NULL, 1);

}
void logout() {
    logflag = 0;
}
CString gendate() {
    char dt[100];
    auto curt=std::chrono::system_clock::now();
    std::time_t curtc = std::chrono::system_clock::to_time_t(curt);
    std::tm curtm;
    //= *std::localtime(&curtc);
    localtime_s(&curtm,&curtc);
    strftime(dt, sizeof(dt), "%d/%m/%Y %X", &curtm);
    CString cdt = CString(dt);
    return cdt;
}
void searchDB(_RecordsetPtr   pRecordset, _ConnectionPtr  pConnection) {
    
    char day[1024], month[1024], year[1024];
    int iser;
    CString searchData[10][100];
    CString cd, cm, cy;
    std::cout << "Enter Day";
    std::cin >> day;
    std::cout << "Enter Month";
    std::cin >> month;
    std::cout << "Enter Year";
    std::cin >> year;
    cd = CString(day);
    cm = CString(month);
    cy = CString(year);
    CString str = "SELECT Count(*) AS Expr2 FROM Dew1 WHERE Date_Accessed LIKE '%" + cd + "/" + cm + "/" + cy + "%';";
    //std::wcout << (LPCTSTR)str;
    pRecordset = pConnection->Execute(_bstr_t(LPCTSTR(str)), NULL, 1);
    iser=_ttoi((LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Expr2")->Value);
    std::cout << iser;
    str="SELECT * FROM Dew1 WHERE Date_Accessed LIKE '%" + cd + "/" + cm + "/" + cy + "%';";
    //std::wcout << (LPCTSTR)str;
    pRecordset=pConnection->Execute(_bstr_t(LPCTSTR(str)), NULL, 1);
    pRecordset->MoveFirst();
    for (int i = 0; i < iser; i++)
    {
        searchData[0][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"ID")->Value;
        searchData[1][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Name")->Value;
        searchData[2][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Image_ID")->Value;
        searchData[3][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Date_Accessed")->Value;
        searchData[4][i] = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"Operations")->Value;
        std::wcout <<(LPCTSTR)searchData[0][i]<<"|"<< (LPCTSTR)searchData[1][i]<<"|"<< (LPCTSTR)searchData[2][i]<<"|"<< (LPCTSTR)searchData[3][i] <<"|" << (LPCTSTR)searchData[4][i] << "\n*--*\n";
        pRecordset->MoveNext();
    }
    
}
void deleteRec(_RecordsetPtr   pRecordset, _ConnectionPtr  pConnection){
    char delid[1024];
    std::cout << "Enter ID of Record to be deleted";
    std::cin >> delid;
    CString cdelid = CString(delid);
    CString query = "DELETE FROM Dew1 WHERE ID=" + cdelid + ";";
    //std::wcout << (LPCTSTR)query;
    pConnection->Execute(_bstr_t(LPCTSTR(query)), NULL, 1);
}
/*
std::vector<char> ReadBMP()
{
    /*int i;
    FILE* f;
   fopen_s(&f, "C:/Users/aayus/Pictures/Screenshots/Model.bmp", "rb");
   unsigned char info[54];
   fread(info, sizeof(unsigned char), 54, f); // read the 54-byte header
   // extract image height and width from header
   int width, height;
   memcpy(&width, info + 18, sizeof(int));
   memcpy(&height, info + 22, sizeof(int));
   int size = 3 * width * abs(height);
   unsigned char* data = new unsigned char[size]; // allocate 3 bytes per pixel
   fread(data, sizeof(unsigned char), size, f); // read the rest of the data at once
   fclose(f);
    return data;
    std::fstream f;
    f.open("C:/Users/aayus/Pictures/Screenshots/BasicView - Copy.bmp", std::ios::in | std::ios::binary);
    const int fileHeaderSize = 14;
    const int informationHeaderSize = 40;
    unsigned char fileHeader[fileHeaderSize];
    f.read(reinterpret_cast<char*>(fileHeader), fileHeaderSize);
    unsigned char informationHeader[informationHeaderSize];
    f.read(reinterpret_cast<char*>(informationHeader), informationHeaderSize);
    int fileSize = fileHeader[2] + (fileHeader[3] << 8) + (fileHeader[4] << 16) + (fileHeader[5] << 24);
    int width = informationHeader[4] + (informationHeader[5] << 8) + (informationHeader[6] << 16) + (informationHeader[7] << 24);
    m_width = width;
    int height = informationHeader[8] + (informationHeader[9] << 8) + (informationHeader[10] << 16) + (informationHeader[11] << 24);
    m_height = height;
    std::vector<Color> m_colors;
    std::vector<char>  bmpdata;
    m_colors.resize(width * height);
    bmpdata.resize(width * height*3);
    const int paddingamount=((4-(width*3)%4)%4);
    for (int y = 0; y < height*3; y++)
    {
        for (int x=0; x <width; x++)
        {
            //unsigned char color[3];
            //f.read(reinterpret_cast<char*>(color), 3);  
            //m_colors[y * width + x].r = static_cast<float>(color[2]) / 255.0f;
            //m_colors[y * width + x].g = static_cast<float>(color[1]) / 255.0f;
            //m_colors[y * width + x].b = static_cast<float>(color[0]) / 255.0f;
            //std::cout<<m_colors[y * width + x].r<<m_colors[y * width + x].g<<m_colors[y * width + x].b;
            bmpdata[y * width + x] = f.get();
            //f.get();
            //f.get();
            std::cout << bmpdata[y * width + x];
        }
        f.ignore(paddingamount);
    }
    f.close();
    std::cout << "file read" << std::endl;
    return bmpdata;
}
std::vector<Color> fetchOLE(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection)
{
    CString cimax = sData[0][imax - 1];
    CString str = "SELECT Image AS img FROM Dew1 WHERE ID="+cimax+";";
    CString data;
    pRecordset = pConnection->Execute(_bstr_t(LPCTSTR(str)), NULL, 1);
    data = (LPCTSTR)(_bstr_t)pRecordset->Fields->GetItem(L"img")->Value;
    int len = data.GetLength();
    std::wcout << (_bstr_t)(LPCTSTR)data;
    std::vector<Color> m_colors;
    m_colors.resize(len);
    for (int i = 0,j=0; i < len-2; i=i+3,j++) {
        if (data[i] == '0')
        {
            m_colors[j].b = static_cast<float>(0) / 255.0f;
        }
        else {
            m_colors[j].b = static_cast<float>(data[i]) / 255.0f;
        }
        if (data[i + 1] == '0')
        {
            m_colors[j].g = static_cast<float>(0) / 255.0f;
        }
        else {
            m_colors[j].g = static_cast<float>(data[i + 1]) / 255.0f;
        }
        if (data[i + 2] == '0') {
            m_colors[j].r = static_cast<float>(0) / 255.0f;
        }
        else {
            m_colors[j].r = static_cast<float>(data[i + 2]) / 255.0f;
        }
        std::cout << m_colors[j].b << "/" << m_colors[j].g << "/" << m_colors[j].r << "\n";
    }
    return m_colors;
}
void exportBMP(const char* path,std::vector<Color> m_colors)
{
    std::fstream f;
    f.open(path, std::ios::out, std::ios::binary);
    unsigned char bmpPad[3] = { 0,0,0 };
    const int paddingAmount = ((4 - (m_width * 3) % 4) % 4);
    std::cout << paddingAmount;
    //_getch();
    const int fileHeaderSize = 14;
    const int informationHeaderSize = 40;
    const int fileSize = fileHeaderSize + informationHeaderSize + m_width * m_height * 3 + paddingAmount * m_width;
    unsigned char fileHeader[fileHeaderSize];
    //file type
    fileHeader[0] = 'B';
    fileHeader[1] = 'M';
    //file size
    fileHeader[2] = fileSize;
    fileHeader[3] = fileSize >> 8;
    fileHeader[4] = fileSize >> 16;
    fileHeader[5] = fileSize >> 24;
    //reserved 1
    fileHeader[6] = 0;
    fileHeader[7] = 0;
    //reserved 2
    fileHeader[8] = 0;
    fileHeader[9] = 0;
    //Pixel data offset
    fileHeader[10] = fileHeaderSize + informationHeaderSize;
    fileHeader[11] = 0;
    fileHeader[12] = 0;
    fileHeader[13] = 0;
    unsigned char informationHeader[informationHeaderSize];
    informationHeader[0] = informationHeaderSize;
    informationHeader[1] = 0;
    informationHeader[2] = 0;
    informationHeader[3] = 0;
    //image width
    informationHeader[4] = m_width;
    informationHeader[5] = m_width>>8;
    informationHeader[6] = m_width>>16;
    informationHeader[7] = m_width>>24;
    informationHeader[8] = m_height;
    informationHeader[9] = m_height>>8;
    informationHeader[10] = m_height>>16;
    informationHeader[11] = m_height>>24;
    //planes
    informationHeader[12] = 1;
    informationHeader[13] = 0;
    //bits per pixel
    informationHeader[14] = 24;
    informationHeader[15] = 0;
    //Compression
    informationHeader[16] = 0;
    informationHeader[17] = 0;
    informationHeader[18] = 0;
    informationHeader[19] = 0;
    //Image size
    informationHeader[20] = 0;
    informationHeader[21] = 0;
    informationHeader[22] = 0;
    informationHeader[23] = 0;
    informationHeader[24] = 0;
    informationHeader[25] = 0;
    informationHeader[26] = 0;
    informationHeader[27] = 0;
    informationHeader[28] = 0;
    informationHeader[29] = 0;
    informationHeader[30] = 0;
    informationHeader[31] = 0;
    informationHeader[32] = 0;
    informationHeader[33] = 0;
    informationHeader[34] = 0;
    informationHeader[35] = 0;
    informationHeader[36] = 0;
    informationHeader[37] = 0;
    informationHeader[38] = 0;
    informationHeader[39] = 0;
    f.write(reinterpret_cast<char*>(fileHeader), fileHeaderSize);
    f.write(reinterpret_cast<char*>(informationHeader), informationHeaderSize);
    for (int y = 0; y < m_height; y++)
    {
        for (int x = 0; x < m_width; x++)
        {
                unsigned char r = static_cast<unsigned char>(m_colors[y * m_width + x].r * 255.0f);
                unsigned char g = static_cast<unsigned char>(m_colors[y * m_width + x].g * 255.0f);
                unsigned char b = static_cast<unsigned char>(m_colors[y * m_width + x].b * 255.0f);
                unsigned char color[] = { b,g,r };
                std::cout << b << g << r;

                f.write(reinterpret_cast<char*>(color), 3);
        }
        f.write(reinterpret_cast<char*>(bmpPad), paddingAmount);

    }
    f.close();
    std::cout << "file created";
}*/
unsigned char* magicFunc()
{

    //cv::Mat image;
    image = cv::imread("C:/Users/aayus/Pictures/Screenshots/URLs.png", cv::IMREAD_UNCHANGED);// Read the file
    std::cout << image.channels()<<"\n";
    if (!image.data) // Check for invalid input
    {
        std::cout << "Could not open or find the image" << std::endl;
    }
    unsigned char* imgdata = image.data;
    std::cout <<"\n IMGDATA\n"<< imgdata;
    cv::Mat TempMat = cv::Mat(image.rows, image.cols, image.type(), imgdata);
    cv::namedWindow("Display window", cv::WINDOW_AUTOSIZE); // Create a window for display.
    imshow("Display window", TempMat); // Show our image inside it.
    cv::waitKey(0);
    return imgdata;// Wait for a keystroke in the window
}
void magicDisplay(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection)
{
    CString cimax = sData[0][imax - 1];
    CString str = "SELECT Image AS img FROM Dew1 WHERE ID=" + cimax + ";";
    CString data;
    pRecordset = pConnection->Execute(_bstr_t(LPCTSTR(str)), NULL, 1);
    data = pRecordset->Fields->GetItem(L"img")->Value;
    CStringA ucStr(data);
    unsigned char* udata = reinterpret_cast<unsigned char*> (ucStr.GetBuffer());
    //data.ReleaseBuffer();
    cv::Mat TempMat = cv::Mat(image.rows, image.cols, image.type(),udata);
    //std::cout << udata;
    cv::imwrite("img2.bmp", TempMat);
    cv::namedWindow("Display window", cv::WINDOW_AUTOSIZE); // Create a window for display.
    imshow("Display window", TempMat); // Show our image inside it.
    cv::waitKey(0);
    //std::wcout << data;

}
//old functions
/*void insImg(CString sData[][100], int imax, _RecordsetPtr   pRecordset, _ConnectionPtr  pConnection, _StreamPtr pStream) {
    //CString data = ReadBMP();
    CString data = "abc";
    CString simax=sData[0][imax - 1];
    CString query = "UPDATE Dew1 SET Image ='"+data+"' WHERE ID ="+simax+";";
    //CString query = "INSERT INTO Dew1 (Image) VALUES('" + data + "') WHERE ID="+simax+"; ";
    //CString cquery(query.c_str());
    std::wcout <<"\n" <<(_bstr_t)(LPCTSTR)query<<"\n";
    pConnection->Execute((_bstr_t)(LPCTSTR)query, NULL, 1);
}*/
//Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\aayus\Documents\Database1.accdb; Persist Security Info=False;

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
