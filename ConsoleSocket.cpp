
#include <iostream>
#include <string>
#include <fstream>
#include <vector>
#include <time.h>

#include <sstream>
#include <iostream>

#pragma warning(disable:4996) 


//#include <winsock.h>
#include <winsock2.h>
#include <ws2tcpip.h>//InetPton

#pragma comment(lib, "Ws2_32.lib")


#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\MSO.DLL"
#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB" 
#import "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" \
    rename("DialogBox","_DialogBox") \
    rename("RGB","_RGB") \
    exclude("IFont","IPicture")

using namespace Excel;


//int sendall(int s, char* buf, int* len)
void sendall(int socket, std::string str)
{
    char* buf = str.data();
    int loc_len= str.size();
    int* len = &loc_len;

    int total = 0; // сколько байт мы послали
    int bytesleft = loc_len; // сколько байт осталось послать
    int n;
    while (total < loc_len) {
        n = send(socket, buf + total, bytesleft, 0);

        std::cout << "############## : "  << n<<std::endl;
        std::cout << "Send:" << buf <<std::endl;
        std::cout << "##############" << std::endl;

        if (n == -1) { break; }
        total += n;
        bytesleft -= n;
    }
    if (loc_len != total) {
        std::cout << "Not full send."<< std::endl;
    }
        
    if (n == -1){
        std::cout << "Send error: -1."<< std::endl;
    }
}

void myLog(std::string text) {

    std::ofstream outfile;
    outfile.open("E:/log.txt", std::ios_base::app);
    time_t     now = time(0);
    struct tm  tstruct;
    char       buf[80];
    tstruct = *localtime(&now);
    strftime(buf, sizeof(buf), "%Y.%m.%d %H:%M:%S", &tstruct);
    std::string String_now = buf;
    outfile << String_now << " | " << text << std::endl;
    outfile.close();

}



int main()
{
    ::CoInitialize(NULL);

    Excel::_ApplicationPtr app("Excel.Application");
    app->Visible[0] = FALSE;
    Excel::_WorkbookPtr book = app->Workbooks->Add();
    Excel::_WorksheetPtr sheet = book->Worksheets->Item[1];


    WSADATA wsaData;
    int iResult = WSAStartup(MAKEWORD(2, 2), &wsaData);
    if (iResult != NO_ERROR) {
        wprintf(L"WSAStartup function failed with error: %d\n", iResult);
        return 1;
    }

    SOCKET ConnectSocket;
    ConnectSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (ConnectSocket == INVALID_SOCKET) {
        wprintf(L"socket function failed with error: %ld\n", WSAGetLastError());
        WSACleanup();
        return 1;
    }


    sockaddr_in clientService;
    memset(&clientService, 0, sizeof clientService);
    clientService.sin_family = AF_INET;
    clientService.sin_addr.s_addr = inet_addr("192.168.1.2");
    clientService.sin_port = htons(7777);
    
    //InetPton(AF_INET, L"192.168.1.2", &clientService.sin_addr.s_addr);
    //clientService.sin_port = htons(7777);

    // Connect to server.
    iResult = connect(ConnectSocket, (SOCKADDR*)&clientService, sizeof(clientService));
    if (iResult == SOCKET_ERROR) {
        wprintf(L"connect function failed with error: %ld\n", WSAGetLastError());
        iResult = closesocket(ConnectSocket);
        if (iResult == SOCKET_ERROR)
            wprintf(L"closesocket function failed with error: %ld\n", WSAGetLastError());
        WSACleanup();
        return 1;
    }

    /*
    std::string data = "Hi atc";
    int len = data.size()+1;
    int bytes_sent = send(ConnectSocket, data.data(), len, 0);
    std::cout << "Send: "<< bytes_sent <<std::endl;

    const int lenBuf = 400;
    char buf[lenBuf];
    memset(buf, 0, lenBuf);

    int get= recv(ConnectSocket, buf, lenBuf, 0);
    std::cout << "Get: " << get << std::endl;

    std::cout << buf << std::endl;
    */
    
    
    sendall(ConnectSocket, "Action: login\r\nUsername: teleami1\r\nSecret: teleami\r\n\r\n");

    /*
    std::vector<std::string> Messages;
    for(int counter_toTest=0; counter_toTest <=100;++counter_toTest){
        std::cout << "### i: " << counter_toTest << std::endl;

        fd_set readfds;
        FD_ZERO(&readfds);
        FD_SET(ConnectSocket, &readfds);
        struct timeval timeout;
        timeout.tv_sec = 5;
        timeout.tv_usec = 0;

        int result_select=select(ConnectSocket + 1, &readfds, NULL, NULL,&timeout);

        if (FD_ISSET(ConnectSocket, &readfds)) {
            const int lenBuf = 1024;
            char buf[lenBuf];
            memset(buf, 0, lenBuf);
            int get = recv(ConnectSocket, buf, lenBuf-1, 0);
            buf[get] = '\0';
            
            std::cout << "Get: " << get << std::endl;
            std::cout << "******************"  << std::endl;
            std::cout << buf << std::endl;
            
            std::string fullMessages = "";
            for (int i_buf = 0; i_buf <= (lenBuf - 3); ++i_buf) {//-1 is my \0 -1 is index not number
                fullMessages.push_back(buf[i_buf]);
            }
            
            if (buf[get - 1] == '\n' && buf[get - 2] == '\r' && buf[get - 3] == '\n' && buf[get - 4] == '\r') {
                std::cout << "Found end." << std::endl;
                Messages.push_back(fullMessages);
                fullMessages = "";
            }
 
            std::cout << "******************" << std::endl;
        }
        
    }
    */

    WSACleanup();

    
    int row = 2;
    int col = 1;
    sheet->Cells->Item[row][col] = "Vasa";
    
    app->Visible[0] = TRUE;
    std::cout << "End main." << std::endl;
}