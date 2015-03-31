/*++

Copyright (c) 2000  Microsoft Corporation

Module Name:

    usbusr.h

Abstract:

Environment:

    Kernel mode

Notes:

    Copyright (c) 2000 Microsoft Corporation.  
    All Rights Reserved.

--*/

#ifndef _USB_USER_H
#define _USB_USER_H

#include <initguid.h>

#ifndef _DEVIOCTL_
#define _DEVIOCTL_

// begin_ntddk begin_wdm begin_nthal begin_ntifs
//
// Define the various device type values.  Note that values used by Microsoft
// Corporation are in the range 0-32767, and 32768-65535 are reserved for use
// by customers.
//

#define DEVICE_TYPE DWORD

#define FILE_DEVICE_BEEP                0x00000001
#define FILE_DEVICE_CD_ROM              0x00000002
#define FILE_DEVICE_CD_ROM_FILE_SYSTEM  0x00000003
#define FILE_DEVICE_CONTROLLER          0x00000004
#define FILE_DEVICE_DATALINK            0x00000005
#define FILE_DEVICE_DFS                 0x00000006
#define FILE_DEVICE_DISK                0x00000007
#define FILE_DEVICE_DISK_FILE_SYSTEM    0x00000008
#define FILE_DEVICE_FILE_SYSTEM         0x00000009
#define FILE_DEVICE_INPORT_PORT         0x0000000a
#define FILE_DEVICE_KEYBOARD            0x0000000b
#define FILE_DEVICE_MAILSLOT            0x0000000c
#define FILE_DEVICE_MIDI_IN             0x0000000d
#define FILE_DEVICE_MIDI_OUT            0x0000000e
#define FILE_DEVICE_MOUSE               0x0000000f
#define FILE_DEVICE_MULTI_UNC_PROVIDER  0x00000010
#define FILE_DEVICE_NAMED_PIPE          0x00000011
#define FILE_DEVICE_NETWORK             0x00000012
#define FILE_DEVICE_NETWORK_BROWSER     0x00000013
#define FILE_DEVICE_NETWORK_FILE_SYSTEM 0x00000014
#define FILE_DEVICE_NULL                0x00000015
#define FILE_DEVICE_PARALLEL_PORT       0x00000016
#define FILE_DEVICE_PHYSICAL_NETCARD    0x00000017
#define FILE_DEVICE_PRINTER             0x00000018
#define FILE_DEVICE_SCANNER             0x00000019
#define FILE_DEVICE_SERIAL_MOUSE_PORT   0x0000001a
#define FILE_DEVICE_SERIAL_PORT         0x0000001b
#define FILE_DEVICE_SCREEN              0x0000001c
#define FILE_DEVICE_SOUND               0x0000001d
#define FILE_DEVICE_STREAMS             0x0000001e
#define FILE_DEVICE_TAPE                0x0000001f
#define FILE_DEVICE_TAPE_FILE_SYSTEM    0x00000020
#define FILE_DEVICE_TRANSPORT           0x00000021
#define FILE_DEVICE_UNKNOWN             0x00000022
#define FILE_DEVICE_VIDEO               0x00000023
#define FILE_DEVICE_VIRTUAL_DISK        0x00000024
#define FILE_DEVICE_WAVE_IN             0x00000025
#define FILE_DEVICE_WAVE_OUT            0x00000026
#define FILE_DEVICE_8042_PORT           0x00000027
#define FILE_DEVICE_NETWORK_REDIRECTOR  0x00000028
#define FILE_DEVICE_BATTERY             0x00000029
#define FILE_DEVICE_BUS_EXTENDER        0x0000002a
#define FILE_DEVICE_MODEM               0x0000002b
#define FILE_DEVICE_VDM                 0x0000002c
#define FILE_DEVICE_MASS_STORAGE        0x0000002d
#define FILE_DEVICE_SMB                 0x0000002e
#define FILE_DEVICE_KS                  0x0000002f
#define FILE_DEVICE_CHANGER             0x00000030
#define FILE_DEVICE_SMARTCARD           0x00000031
#define FILE_DEVICE_ACPI                0x00000032
#define FILE_DEVICE_DVD                 0x00000033
#define FILE_DEVICE_FULLSCREEN_VIDEO    0x00000034
#define FILE_DEVICE_DFS_FILE_SYSTEM     0x00000035
#define FILE_DEVICE_DFS_VOLUME          0x00000036

//
// Macro definition for defining IOCTL and FSCTL function control codes.  Note
// that function codes 0-2047 are reserved for Microsoft Corporation, and
// 2048-4095 are reserved for customers.
//

#define CTL_CODE( DeviceType, Function, Method, Access ) (                 \
    ((DeviceType) << 16) | ((Access) << 14) | ((Function) << 2) | (Method) \
)
//
// Define the method codes for how buffers are passed for I/O and FS controls
//

#define METHOD_BUFFERED                 0
#define METHOD_IN_DIRECT                1
#define METHOD_OUT_DIRECT               2
#define METHOD_NEITHER                  3

//
// Define the access check value for any access
//
//
// The FILE_READ_ACCESS and FILE_WRITE_ACCESS constants are also defined in
// ntioapi.h as FILE_READ_DATA and FILE_WRITE_DATA. The values for these
// constants *MUST* always be in sync.
//


#define FILE_ANY_ACCESS                 0
#define FILE_READ_ACCESS          ( 0x0001 )    // file & pipe
#define FILE_WRITE_ACCESS         ( 0x0002 )    // file & pipe

// end_ntddk end_wdm end_nthal end_ntifs

#endif // _DEVIOCTL_


/*#define USB_IOCTL_INDEX             0x0000


#define IOCTL_USB_GET_CONFIG_DESCRIPTOR CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)
                                                   
#define IOCTL_USB_RESET_DEVICE          CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 1, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_RESET_PIPE            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 2, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_VENDOR_REQUEST CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 3,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_ANCHOR_DOWNLOAD CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 4,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB1_ANCHOR_DOWNLOAD CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 5,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_CURRENT_CONFIG CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 6,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_CURRENT_FRAME_NUMBER CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 7,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_RESETPIPE CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 8,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_ABORTPIPE CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 9,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_PIPE_INFO CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 10,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_DEVICE_DESCRIPTOR CTL_CODE(FILE_DEVICE_KEYBOARD,     \
                                                     USB_IOCTL_INDEX + 0,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_STRING_DESCRIPTOR CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 12,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_CONFIGURATION_DESCRIPTOR CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 13,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_SETINTERFACE CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 14,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_RESET CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 15,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_BULK_WRITE CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 16,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_BULK_READ CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 17,     \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_VENDOR_OR_CLASS_REQUEST            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 18, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_LAST_ERROR            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 19, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_ISO_READ            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 20, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_ISO_WRITE            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 21, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_START_ISO_STREAM           CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 22, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_STOP_ISO_STREAM            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 23, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_READ_ISO_BUFFER            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 24, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_GET_DRIVER_VERSION           CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 25, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_SET_FEATURE            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 26, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_LOOP_BACK            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 27, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)

#define IOCTL_USB_WAIT_INT            CTL_CODE(FILE_DEVICE_UNKNOWN,     \
                                                     USB_IOCTL_INDEX + 28, \
                                                     METHOD_BUFFERED,         \
                                                     FILE_ANY_ACCESS)*/

#endif

