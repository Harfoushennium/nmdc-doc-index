// Research-only native worksheet structural scanner. Expat supplies XML decoding.
#include <expat.h>
#include <algorithm>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <sstream>
#include <fstream>
#include <iostream>
#include <stdexcept>
#include <string>
#include <vector>
using namespace std;
static const string NS="http://schemas.openxmlformats.org/spreadsheetml/2006/main|";
static const string REL="http://schemas.openxmlformats.org/officeDocument/2006/relationships|id";
struct Cell {string ref,type,value,formula,inline_text;};
struct Row {string number; vector<Cell> cells;};
struct Scan {
 vector<Row> rows; vector<string> merges; vector<vector<string>> links;
 string dimension; uint64_t hidden_rows=0,hidden_columns=0;
 vector<string> stack; bool active_row=false,active_cell=false; Row row; Cell cell;
 bool failed=false; XML_Parser parser=nullptr;
};
string attr(const XML_Char** a,const string& k){for(;a&&*a;a+=2)if(k==a[0])return a[1];return "";}
bool hidden(const string& s){return s=="1"||s=="true"||s=="TRUE";}
static void start(void* ptr,const XML_Char* name,const XML_Char** a){
 auto& s=*static_cast<Scan*>(ptr);
 try{
 string n=name; s.stack.push_back(n);
 if(n==NS+"dimension" && s.dimension.empty())s.dimension=attr(a,"ref");
 else if(n==NS+"row"){s.active_row=true;s.row=Row{attr(a,"r"),{}};if(hidden(attr(a,"hidden")))s.hidden_rows++;}
 else if(n==NS+"col"&&hidden(attr(a,"hidden")))s.hidden_columns++;
 else if(n==NS+"c"&&s.active_row){s.active_cell=true;s.cell=Cell{attr(a,"r"),attr(a,"t"),"","",""};}
 else if(n==NS+"mergeCell")s.merges.push_back(attr(a,"ref"));
 else if(n==NS+"hyperlink")s.links.push_back({attr(a,"ref"),attr(a,REL),attr(a,"location")});
 }catch(...){s.failed=true;XML_StopParser(s.parser,XML_FALSE);}
}
static void chars(void* ptr,const XML_Char* data,int len){
 auto& s=*static_cast<Scan*>(ptr);if(!s.active_cell||s.stack.empty())return;
 try{auto& n=s.stack.back();if(n==NS+"v")s.cell.value.append(data,len);else if(n==NS+"f")s.cell.formula.append(data,len);else if(n==NS+"t")s.cell.inline_text.append(data,len);}
 catch(...){s.failed=true;XML_StopParser(s.parser,XML_FALSE);}
}
static void end(void* ptr,const XML_Char* name){
 auto& s=*static_cast<Scan*>(ptr);
 try{string n=name;if(n==NS+"c"&&s.active_cell){s.row.cells.push_back(move(s.cell));s.active_cell=false;}
 else if(n==NS+"row"&&s.active_row){s.rows.push_back(move(s.row));s.active_row=false;}s.stack.pop_back();}
 catch(...){s.failed=true;XML_StopParser(s.parser,XML_FALSE);}
}
static void doctype(void* ptr,const XML_Char*,const XML_Char*,const XML_Char*,int){
 auto& s=*static_cast<Scan*>(ptr);s.failed=true;XML_StopParser(s.parser,XML_FALSE);
}
void field(ostream& o,const string& s){if(s.size()>0xffffffffu)throw runtime_error("oversize field");uint32_t n=(uint32_t)s.size();for(int i=0;i<4;i++)o.put(char((n>>(8*i))&255));o.write(s.data(),s.size());}
void number(ostream& o,uint64_t n){field(o,to_string(n));}
static bool parse(istream& in,Scan& s){
 s.parser=XML_ParserCreateNS(nullptr,'|');if(!s.parser)return false;
 XML_SetUserData(s.parser,&s);XML_SetElementHandler(s.parser,start,end);XML_SetCharacterDataHandler(s.parser,chars);XML_SetStartDoctypeDeclHandler(s.parser,doctype);
 char buffer[65536];bool ok=true;
 do{in.read(buffer,sizeof(buffer));auto count=in.gcount();if(XML_Parse(s.parser,buffer,(int)count,in.eof())==XML_STATUS_ERROR){ok=false;break;}}while(!in.eof()&&in);
 XML_ParserFree(s.parser);return ok&&!s.failed&&!in.bad();
}
static void write_scan(ostream& out,Scan& s){
 field(out,"NMDC-SCAN-1");field(out,s.dimension);number(out,s.hidden_rows);number(out,s.hidden_columns);
 number(out,s.rows.size());for(auto&r:s.rows){field(out,r.number);number(out,r.cells.size());for(auto&c:r.cells){field(out,c.ref);field(out,c.type);field(out,c.value);field(out,c.formula);field(out,c.inline_text);}}
 number(out,s.merges.size());for(auto&m:s.merges)field(out,m);
 number(out,s.links.size());for(auto&l:s.links)for(auto&v:l)field(out,v);
 if(!out)throw runtime_error("output failure");
}
extern "C" int nmdc_scan_xml(const char* data,size_t size,void** output,size_t* output_size){
 if(!output||!output_size)return 1;
 *output=nullptr;*output_size=0;
 try{
 if(!data||size>256u*1024u*1024u)return 1;
 istringstream in(string(data,size),ios::binary|ios::in);Scan s;
 if(!parse(in,s))return 2;
 ostringstream out(ios::binary|ios::out);write_scan(out,s);string raw=out.str();
 void* mem=malloc(raw.size());if(!mem)return 3;
 memcpy(mem,raw.data(),raw.size());*output=mem;*output_size=raw.size();return 0;
 }catch(...){return 4;}
}
extern "C" void nmdc_scan_free(void* p){free(p);}
int main(int argc,char**argv){
 try{
 if(argc!=3)throw runtime_error("usage: scan input.xml output.bin");
 ifstream in(argv[1],ios::binary);if(!in)throw runtime_error("input open failed");
 Scan s;if(!parse(in,s))throw runtime_error("invalid or unsupported XML");
 ofstream out(argv[2],ios::binary);if(!out)throw runtime_error("output open failed");
 write_scan(out,s);return 0;
 }catch(const exception&e){cerr<<e.what()<<'\n';return 1;}
}
