/*
(c) Copyright 1983 - 2016 Keysight Technologies Inc.
The contents of this software are proprietary and confidential to Keysight
Technologies, and are limited in distribution to those with a direct need
to know.  Individuals having access to this software are responsible for main-
taining the confidentiality of the content and for keeping the software secure
when not in use.  Transfer to any party is strictly forbidden other than as
expressly permitted in writing by Keysight Technologies.  Unauthorized trans-
fer to or possession by any unauthorized party may be a criminal offense.

RESTRICTED RIGHTS LEGEND

Use,  duplication,  or disclosure by the Government  is
subject to restrictions as set forth in subdivision (b)
(3)  (ii)  of the Rights in Technical Data and Computer
Software clause at 52.227-7013.

KEYSIGHT TECHNOLOGIES
1400 Fountaingrove Parkway
Santa Rosa, CA 95403-1738
United States
*/
#define __STR2__(x) #x 
#define __STR1__(x) __STR2__(x) 
#define __LOC__ __FILE__ "("__STR1__(__LINE__)") : Warning: " 
#pragma message(__LOC__"The file has been deprecated.") 

#if !defined(AGMODINST_H)
#define AGMODINST_H

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers
#include <windows.h>

#ifdef AG_PCIE_MEM_DMA_EXPORTS
#define AG_PCIE_MEM_DMA_API __declspec(dllexport)
#else
#define AG_PCIE_MEM_DMA_API __declspec(dllimport)
#endif

#if defined(_WIN64)
#pragma message ("Note: This is 64-bit build")
typedef unsigned	__int64 IolsPcieLength;
typedef unsigned  __int64 IolsPcieBarOffset;
typedef unsigned  __int64 IolsMapStart;
typedef unsigned  __int64 IolsMapSize;
#else
/* This is a 32-bit OS, not a 64-bit OS */
#pragma message ("Note: This is 32-bit build")
typedef unsigned int  IolsPcieBarOffset;
typedef unsigned int  IolsPcieLength;
typedef unsigned int  IolsMapStart;
typedef unsigned int  IolsMapSize;
#endif

#define USE_DMA_MASK 1
#define USE_OOOWRITES_MASK 2
#define UseDma(x) ((x & USE_DMA_MASK) != 0)
#define UseOooWrites(x) ((x & USE_OOOWRITES_MASK) != 0)
#define PXI_FLAGS_MASK 0x0000FF00

// class UserMap
class agModInstMap
{
public:
   void         *kernelMapHandle;
   int          space;
   IolsMapStart mapStart;
   IolsMapSize  mapSize;
   int          dereferenceable;
   void         *memPointer;
   static void* getMemPointer(agModInstMap *m)
   {
      void *retVal = NULL;
      if (m != NULL)
      {
         retVal = m->memPointer;
      }
      return retVal;
   }
};

typedef enum
{
	Bar0 = 0,
	Bar1,
	Bar2,
	Bar3,
	Bar4,
	Bar5,
   Config,
   Alloc
} BAR;

// A pointer to this struct is used when doing IoControl operations from VISA
// using viSetAttribute(vi, VI_KTATTR_DEVICE_IO_CONTROL, &agIoctlStruct)
typedef struct {
   DWORD        dwStructSize;     // The size of this structure in bytes (used for sanity check)
   DWORD        dwIoControlCode;
   LPVOID       lpInBuffer;
   DWORD        dwInBufferSize;
   LPVOID       lpOutBuffer;
   DWORD        dwOutBufferSize;
   LPDWORD      lpBytesReturned;
   LPOVERLAPPED lpOverlapped;
   LPBOOL       lpReturnCode;     // This is the value returned from the DeviceIoControl call
} deviceIoControlParams;

typedef DWORD (*IOLS_PCIE_OPEN)(int intfc, int bus, int device, int function, HANDLE *pHandle);
typedef DWORD (*IOLS_PCIE_MAP_MEM)(HANDLE h,BAR bar,IolsPcieBarOffset offset,IolsPcieLength length,void **pUserSpaceMem);
typedef DWORD (*IOLS_PCIE_UNMAP_MEM)(HANDLE h,BAR bar,IolsPcieBarOffset offset, IolsPcieLength length,void *pUserSpaceMem);
typedef DWORD (*IOLS_PCIE_IOCTL)(HANDLE h,DWORD dwIoControlCode,LPVOID lpInBuffer,DWORD nInBufferSize,LPVOID lpOutBuffer,DWORD nOutBufferSize,LPDWORD lpBytesReturned,DWORD dwTimeoutMilliseconds);
typedef DWORD (*IOLS_PCIE_BLOCK_WRITE)(HANDLE h,int timeoutMs,int useDma,void *mapHandle,IolsPcieLength offset,int width,bool increment,void *writeBuffer,IolsPcieLength  count);
typedef DWORD (*IOLS_PCIE_BLOCK_READ)(HANDLE h,int timeoutMs,int useDma,void *mapHandle,IolsPcieLength offset,int width,bool increment,void *readBuffer,IolsPcieLength count);
typedef DWORD (*IOLS_PCIE_WAIT_INTERRUPT)(HANDLE h);
typedef DWORD (*IOLS_PCIE_ABORT_WAIT_INTERRUPT)(HANDLE h);
typedef DWORD (*IOLS_PCIE_READ_PCI_CONFIG_SPACE)(HANDLE h, IolsPcieBarOffset offset, int length, void *pBuffer);
typedef DWORD (*IOLS_PCIE_CLOSE)(HANDLE h);
typedef DWORD (*IOLS_PCIE_GET_DEVICE_SLOT_NUMBER)(HANDLE h,int* pslot_num);
typedef DWORD (*IOLS_PCIE_TERMINATE_IO)(HANDLE h, LPVOID data);

extern "C"
{
   // Opens and closes a handle to the Agilent common PXI kernel driver
   int agmiOpen(int interfaceNumber, // IN, must be 0 for normal PXI access
                int bus,             // IN, PCI bus number
                int device,          // IN, PCI device number
                int function,        // IN, PCI function number
                void **pxiHandle     // OUT, handle value returned
                );

   int agmiClose(void *pxiHandle // IN, handle returned from agmiOpen
                 );

   int agmiLock(void *pxiHandle, // IN, handle returned from agmiOpen
                bool wait,       // IN, whether or not to wait to acquire the lock or return immediately 
                int timeoutMs    // IN, the number of milliseconds to wait for the lock before returning
                );

   int agmiUnlock(void *pxiHandle // IN, handle returned from agmiOpen
                  );

   // This function blocks waiting for a PXI interrupt so it should be called on a separate thread.
   // It will return when an interrupt occurs or if agmiAbortWaitInterrupt is called.
   //Legacy: use agmiPpiEnableInterrupts() and agmiPpiWaitInterrupt() for new code.
   int agmiWaitInterrupt(void *pxiHandle // IN, handle returned from agmiOpen
                         );

   //Legacy: use agmiPpiDisableAndAbortWaitInterrupt() for new code.
   int agmiAbortWaitInterrupt(void *pxiHandle // IN, handle returned from agmiOpen
                           );

   // These interrupt control functions mirror those defined in IVI 6-3.
   // They replace agmiWaitInterrupt and agmiAbortWaitInterrupt
   int agmiPpiEnableInterrupts(void *pxiHandle, unsigned short queueLength);
   int agmiPpiWaitInterrupt(void *pxiHandle, unsigned int timeoutMilliseconds, short *interruptSequence, unsigned int *interruptData);
   int agmiPpiDisableAndAbortWaitInterrupt(void *pxiHandle);

   // Map and Unmap memory from a PXI device.

   // Note: The agmiMapMemory function indicates whether the PXI memory can be dereferenced by the user.
   int agmiMapMemory(void         *pxiHandle,      // IN, handle returned from agmiOpen
                     int          addrSpace,       // IN, the PXI memory space to map
                     IolsMapStart offset,          // IN, the offset in bytes for the beginning of the map
                     IolsMapSize  length,          // IN, the length of the map in bytes
                     void         **mapHandle,     // OUT, a handle to the memory map
                     void         **memPtr,        // OUT, a user-space pointer to the memory map if it is dereferencable
                     int          *dereferenceable // OUT, 0 if memory cannont be dereferenced, 1 if it can
                     );

   int agmiUnmapMemory(void *pxiHandle, // IN, handle returned from agmiOpen
                       void *mapHandle  // IN, handle returned from agmiMapMemory
                       );

   int agmiMapMemory2(void             *pxiHandle,      // IN, handle returned from agmiOpen
                      BAR              barSpace,        // IN, the PXI memory space to map
                      unsigned __int64 offset,          // IN, the offset in bytes for the beginning of the map
                      unsigned __int64 length,          // IN, the length of the map in bytes
                      void             **mapHandle2,    // OUT, a handle to the memory map
                      void             **memPtr,        // OUT, a user-space pointer to the memory map if it is dereferencable
                      int              *dereferenceable // OUT, 0 if memory cannont be dereferenced, 1 if it can
                      );

   int agmiUnmapMemory2(void *pxiHandle, // IN, handle returned from agmiOpen
                        void *mapHandle2 // IN, handle returned from agmiMapMemory2
                        );

   int agmiMemAlloc(void *pxiHandle,             // IN, handle returned from agmiOpen
                    unsigned __int64 memSize,    // IN, handle returned from agmiMapMemory
                    unsigned __int64 *busAddress // OUT, address of contiguous locked physical memory
                    );

   int agmiMemFree(void *pxiHandle,             // IN, handle returned from agmiOpen
                    unsigned __int64 busAddress // IN, address of contiguous locked physical memory
                    );

   // Memory access functions
   int agmiPeek(void        *pxiHandle, // IN, handle returned from agmiOpen
                void        *mapHandle, // IN, handle returned from agmiMapMemory
                IolsMapSize offset,     // IN, offset
                int          width,     // IN, data width in bits
                void         *value,    // OUT, where to put peek'd value
                int          timeoutMs  // IN, timeout
                );

   int agmiPeek2(void             *pxiHandle,   // IN, handle returned from pxiOpen
                 void             *mapHandle2,  // IN, handle returned from pxiMapMemory2
                 unsigned __int64 offset,       // IN, offset in bytes
                 int              widthInBytes, // IN, data width in bytes
                 void             *value,       // OUT, where to put peek'd value
                 int              timeoutMs     // IN, timeout
                 );

   int agmiPoke(void             *pxiHandle, // IN, handle returned from agmiOpen
                void             *mapHandle, // IN, handle returned from agmiMapMemory
                IolsMapSize      offset,     // IN, offset
                int              width,      // IN, data width in bits
                unsigned __int64 value,      // IN, value to poke
                int              timeoutMs   // IN, timeout
                );

   int agmiPoke2(void             *pxiHandle,   // IN, handle returned from pxiOpen
                 void             *mapHandle2,  // IN, handle returned from pxiMapMemory2
                 unsigned __int64 offset,       // IN, offset
                 int              widthInBytes, // IN, data width in bytes
                 unsigned __int64 value,        // IN, value to poke
                 int              timeoutMs     // IN, timeout
                 );

   int agmiBlockWrite(void        *pxiHandle,   // IN, handle returned from agmiOpen
                      int         timeoutMs,    // IN, timeout in Milliseconds
                      int         flags,        // IN, use DMA (bit 0) oooWrites (bit 1)
                      void        *mapHandle,   // IN, handle returned from agmiMapMemory
                      IolsMapSize offset,       // IN, byte offset from base of map
                      int         width,        // IN, width (8, 16, 32 or 64)
                      bool        increment,    // IN, increment (0 for fifo, 1 for copy)
                      void        *writeBuffer, // IN, address of memory for data to write
                      IolsMapSize count,        // IN, count (bytes)
                      bool        swap          // IN, swap or not
                      );

   int agmiBlockWrite2(void             *pxiHandle,   // IN, handle returned from agmiOpen
                       int              timeoutMs,    // IN, timeout in Milliseconds
                       int              flags,        // IN, use DMA (bit 0) oooWrites (bit 1)
                       BAR              barSpace,     // IN, BAR enum value
                       unsigned __int64 byteOffset,   // IN, byte offset from base of map
                       int              widthInBytes, // IN, width (8, 16, 32 or 64)
                       bool             increment,    // IN, increment (0 for fifo, 1 for copy)
                       void             *writeBuffer, // IN, address of memory for data to write
                       unsigned __int64 elementCount  // IN, count in elements
                       );

   int agmiBlockRead(void        *pxiHandle,  // IN, handle returned from agmiOpen
                     int         timeoutMs,   // IN, timeout in Milliseconds
                     int         flags,       // IN, use DMA (bit 0)
                     void        *mapHandle,  // IN, Ihandle returned from agmiMapMemory
                     IolsMapSize offset,      // IN, byte offset from base of map
                     int         width,       // IN, width (8, 16, 32 or 64)
                     bool        increment,   // IN, increment (0 for fifo, 1 for copy)
                     void        *readBuffer, // OUT,address of user memory for returned data
                     IolsMapSize count,       // IN, count (bytes)
                     bool        swap         // IN, swap or not
                     );

   int agmiBlockRead2(void             *pxiHandle,   // IN, handle returned from agmiOpen
                      int              timeoutMs,    // IN, timeout in Milliseconds
                      int              flags,        // IN, use DMA (bit 0) oooWrites (bit 1)
                      BAR              barSpace,     // IN, BAR enum value
                      unsigned __int64 byteOffset,   // IN, byte offset from base of map
                      int              widthInBytes, // IN, width (8, 16, 32 or 64)
                      bool             increment,    // IN, increment (0 for fifo, 1 for copy)
                      void             *readBuffer,  // IN, address of user memory for returned data
                      unsigned __int64 elementCount  // IN, count in elements
                      );

   int agmiBlockCopy(void        *pxiHandle,    // IN, handle returned from agmiOpen
                     int         timeoutMs,     // IN, timeout in Milliseconds
                     int         flags,         // IN, use DMA (bit 0) oooWrites (bit 1)
                     void        *srcMapHandle, // IN, source Ihandle returned from agmiMapMemory
                     IolsMapSize srcOffset,     // IN, source byte offset from base of map
                     int         srcWidth,      // IN, source width (8, 16, 32 or 64)
                     bool        srcIncrement,  // IN, source increment (0 for fifo, 1 for copy)
                     void        *dstMapHandle, // IN, destination Ihandle returned from agmiMapMemory
                     IolsMapSize dstOffset,     // IN, destination byte offset from base of map
                     int         dstWidth,      // IN, destination width (8, 16, 32 or 64)
                     bool        dstIncrement,  // IN, destination increment (0 for fifo, 1 for copy)
                     IolsMapSize count,         // IN, count (bytes)
                     bool        swap           // IN, swap or not
                     );

   int agmiDeviceIoControl(void         *pxiHandle,      // IN, handle returned from agmiOpen
                           DWORD        dwIoControlCode, // IN, IOCTL code
                           LPVOID       lpInBuffer,      // IN, pointer to input buffer
                           DWORD        dwInBufferSize,  // IN, length in bytes of input buffer
                           LPVOID       lpOutBuffer,     // OUT, pointer to output buffer
                           DWORD        dwOutBufferSize, // IN, size in bytes of output buffer
                           LPDWORD      lpBytesReturned, // OUT, size of data stored in output buffer
                           LPOVERLAPPED lpOverlapped,    // INOUT, pointer to overlapped structure
                           LPBOOL       lpReturnCode     // OUT, pointer to return code from DeviceIoControl call
                           );

	int agmiSetDetectQuiesce(void *pxiHandle,LPVOID lpInBuffer,int nInBufferSize);
	int agmiSetDriverBehavior( void *pxiHandle,void *data,int len );
	int agmiGetDriverBehavior( void *pxiHandle,void *data,int len );
   int agmiCmd(void *pxiHandle, long command, int arraySize, int elementSize, void *data);
   int agmiTerminateIo(void *pxiHandle, void *data);

   // Low level API functions which do not rely on the Agilent PXI Resource Manager having been run:
   AG_PCIE_MEM_DMA_API DWORD agPcieOpen(int intfc, int bus, int device,int function,HANDLE *pHandle);
   AG_PCIE_MEM_DMA_API DWORD agPcieClose(HANDLE handle);
   AG_PCIE_MEM_DMA_API DWORD agPcieMapMemory(HANDLE handle,BAR bar,IolsMapStart offset,IolsMapSize length,void **pUserSpaceMem);
   AG_PCIE_MEM_DMA_API DWORD agPcieUnmapMemory(HANDLE handle,BAR bar, IolsMapStart offset, IolsMapSize length,void *pUserSpaceMem);
   AG_PCIE_MEM_DMA_API DWORD agPcieDeviceIoControl(HANDLE handle,DWORD dwIoControlCode,LPVOID lpInBuffer,DWORD nInBufferSize,LPVOID lpOutBuffer,DWORD nOutBufferSize,LPDWORD lpBytesReturned,DWORD dwTimeoutMilliseconds, HANDLE hCancelIo=NULL);
   AG_PCIE_MEM_DMA_API DWORD agPcieGetDeviceSlotNumber(HANDLE handle, int* slot_num);
   AG_PCIE_MEM_DMA_API DWORD agPcieSetDetectQuiesce(HANDLE handle,LPVOID lpInBuffer,int nInBufferSize);
   AG_PCIE_MEM_DMA_API DWORD agPcieSetDriverBehavior(HANDLE handle,LPVOID lpInBuffer,int nInBufferSize);
   AG_PCIE_MEM_DMA_API DWORD agPcieGetDriverBehavior(HANDLE handle,LPVOID lpInBuffer,int nInBufferSize);
   AG_PCIE_MEM_DMA_API DWORD agPcieReadCfgSpace(HANDLE handle, IolsPcieBarOffset offset, int length, void *pBuffer);
   AG_PCIE_MEM_DMA_API DWORD agPcieTerminateIo(HANDLE handle, HANDLE hCancelEvent, void *data);
   AG_PCIE_MEM_DMA_API DWORD agPcieGetSpaceInfo(HANDLE handle, BAR bar, short *spaceType, unsigned __int64 *spaceBase, unsigned __int64 *spaceSize);
   AG_PCIE_MEM_DMA_API DWORD agPcieGetDeviceAttribute(HANDLE handle, UINT32 attributeID, void * attributeValue);

   // These functions now obsolete.  The corresponding agPcie... calls can now handle translation DLL calls transparently.
   // They are left here for backwards compatibility.
   AG_PCIE_MEM_DMA_API DWORD agPcieX_Open(int intfc, int bus, int device,int function,HANDLE *pHandle);
   AG_PCIE_MEM_DMA_API DWORD agPcieX_Close(HANDLE handle);
   AG_PCIE_MEM_DMA_API DWORD agPcieX_ReadCfgSpace(HANDLE h, IolsPcieBarOffset offset, int length, void *pBuffer);
   AG_PCIE_MEM_DMA_API DWORD agPcieX_GetDeviceSlotNumber(HANDLE h, int* pslot_num);
} // extern "C"

#endif // AGMODINST_H
