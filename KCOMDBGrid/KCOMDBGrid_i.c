/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Fri May 07 23:33:08 1999
 */
/* Compiler settings for D:\wxz\KCOMDBGrid\KCOMDBGrid.odl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID LIBID_KCOMDBGRIDLib = {0xAC212643,0xFBE2,0x11D2,{0xA7,0xFE,0x00,0x80,0xC8,0x76,0x3F,0xA4}};


const IID DIID__DKCOMDBGrid = {0xAC212644,0xFBE2,0x11D2,{0xA7,0xFE,0x00,0x80,0xC8,0x76,0x3F,0xA4}};


const IID DIID__DKCOMDBGridEvents = {0xAC212645,0xFBE2,0x11D2,{0xA7,0xFE,0x00,0x80,0xC8,0x76,0x3F,0xA4}};


const CLSID CLSID_KCOMDBGrid = {0xAC212646,0xFBE2,0x11D2,{0xA7,0xFE,0x00,0x80,0xC8,0x76,0x3F,0xA4}};


#ifdef __cplusplus
}
#endif

