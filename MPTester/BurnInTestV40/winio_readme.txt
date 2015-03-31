A portion of Yariv Kaplan's WinIo library was used in this product.
As requested the legal notice is reproduced here.


------------------------------------------------------------
                        WinIo v1.2                          
    Direct Hardware Access Under Windows 9x/NT/2000         
            Copyright 1998-2000 Yariv Kaplan                
                http://www.internals.com                    
------------------------------------------------------------

The WinIo library allows 32-bit Windows applications to directly
access I/O ports and physical memory. It bypasses Windows protection
mechanisms by using a combination of a kernel-mode device driver and
several low-level programming techniques.

Under Windows NT, the WinIo library can only be used by applications
that have administrative privileges. If the user is not logged on as
an administrator, the WinIo DLL is unable to install and activate the
WinIo driver. It is possible to overcome this limitation by installing
the driver once through an administrative account. In that case, however,
the ShutdownWinIo function must not be called before the application
is terminated, since this function removes the WinIo driver from the
system's registry.

----------------------------------------------------------------------------
                              LEGAL STUFF             
----------------------------------------------------------------------------

The following terms apply to all files associated with the software
unless explicitly disclaimed in individual files.

The author hereby grants permission to use, copy, modify, distribute,
and license this software and its documentation for any purpose, provided
that existing copyright notices are retained in all copies and that this
notice is included verbatim in any distributions. No written agreement,
license, or royalty fee is required for any of the authorized uses.

IN NO EVENT SHALL THE AUTHOR OR DISTRIBUTORS BE LIABLE TO ANY PARTY
FOR DIRECT, INDIRECT, SPECIAL, INCIDENTAL, OR CONSEQUENTIAL DAMAGES
ARISING OUT OF THE USE OF THIS SOFTWARE, ITS DOCUMENTATION, OR ANY
DERIVATIVES THEREOF, EVEN IF THE AUTHOR HAVE BEEN ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.

THE AUTHOR AND DISTRIBUTORS SPECIFICALLY DISCLAIM ANY WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT. THIS SOFTWARE
IS PROVIDED ON AN "AS IS" BASIS, AND THE AUTHOR AND DISTRIBUTORS HAVE
NO OBLIGATION TO PROVIDE MAINTENANCE, SUPPORT, UPDATES, ENHANCEMENTS, OR
MODIFICATIONS.


I can be reached at: yariv@internals.com

Yariv Kaplan
