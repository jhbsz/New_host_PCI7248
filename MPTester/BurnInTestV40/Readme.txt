PassMark BurnInTest V4.0
Copyright (C) 2004 PassMark Software
All Rights Reserved
http://www.passmark.com

Overview
========
Passmark's BurnInTest is a software tool that allows all
the major sub-systems of a computer to be simultaneously
tested for reliability and stability.
<For more details see the online help>


Status
======
This is a shareware program.
This means that you need to buy it if you would like
to continue using it after the evaluation period.

Installation
============
1) Uninstall any previous version of BurnInTest
2) Double click (or Open) the downloaded ".exe" file
3) Follow the prompts


UnInstallation
==============
Use the Windows control panel, Add / Remove Programs


Requirements
============
- Operating System: Windows 98, 2000, ME, XP, 2003 server (*)
- RAM: 32 Meg 
- Disk space: 2 Meg of free hard disk space (plus an additional 10Meg to run 
  the Disk test)
- DirectX 8 or above software for 3D graphics and video tests (plus working 
  DirectX drivers for your video card)
- MMX compatible CPU for MMX tests
- A printer to run the printer test, set-up as the default printer in Windows.
- A CD ROM + 1 Music CD or Data CD to run the CD test.
- A CD-RW to run the CD burn test.
- A network connection and the TCP/IP networking software installed for the 
  Network Tests
- A serial port loop back plug for the serial port test. (Pro version only)
- A parallel port loop back plug for the parallel port test. (Pro version only)
- A USB port loop back plug for the USB port test. (Pro version only)
- A USB 2.0 port loop back plug for the USB 2.0 port test. (Pro version only)
(*) Windows98 and Windows Me do not support the Tape drive, Video playback, 
CD-RW burn, USB2 or the Disk test mode of Butterfly seeking tests. Windows 
2000 does not support the CD-RW burn test. The advanced RAM test is only 
available under Windows 2000 and Windows XP professional (the other RAM tests 
are supported under the other OS's). Users must have administrator privileges 
in 2000 and XP.

Windows 95 and Windows NT
=========================
Windows 95 and Windows NT are not supported in BurnInTest version 4.0 and above. 
Use a version of BurnInTest prior to 3.1 for compatibility with W95 and NT.


Version History
===============
Here is a summary of all changes that have been made in each version of 
BurnInTest.

Release 4.0 build 1011, 1/April/2004.

- Minor changes of the USB Loopback test. These relate to the upgrade of the 
  USB Loopback plug (blue plug) device drivers to V2.0.0.0. Improvements have 
  also been made with the error handling after a USB loopback plug is disconnected 
  during a test.
  Note: It is recommended that V4.0 build 1011 or above of BurnInTest be used with 
  V2.0.0.0 of the USB loopback plug device drivers.

Release 4.0 build 1010, 18/March/2004.

- Change to allow the registration of a valid username/key even when the system
  clock setting is incorrect.
- Correction to resizing of the Log Details window.
- Correction to reported error count for the random seeking test mode of the disk 
  test. Increased activity level 2 tracing for the random seek test mode.

Release 4.0 build 1009, 25/February/2004

- Rare USB2.0 loopback test bug corrected.

Release 4.0 build 1008, 23/February/2004

- The disk drive Butterfly seeking test mode requires the disk device driver to 
  support the control code for supplying DISK_GEOMETRY_EX information. Some disks 
  do not support this control code and the older control code is now also attempted.
  If both fail, a new error message is displayed indicating that the Butterfly
  test is not supported by this disk drive. This replaces the error "Unable to 
  get disk geometry". This 'INFORMATION' level error may be ignored by editing 
  the BITClassification.txt file.

Release 4.0 build 1007, 18/February/2004

- Corrected a bug where an old version of sensorDLL.dll on some Intel Motherboard PC's 
  would crash BurnInTest on startup.
- Changed the number of 2D graphics operations per cycle to 20%  of previous the 
  previous value.
- Changed the number of Parallel port operations per cycle to 60% of previous the 
  previous value.
- Corrected a bug that only allowed 19 hard disks to be tested, instead of 20.

Release 4.0 build 1006, 6/February/2004

- Text of Log Clearing options in Logging preferences changed to better reflect
  functionality. Option added: Clear test results and create new log.
- Improved the USB test error handling such that when a USB loopback plug or a 
  USB 2.0 loopback plug is unplugged and re-plugged during a test, automatic device 
  detection is attempted and allows the test to continue when possible.
- Stop button removed from being on top of the EMC Video Display Unit test.

Release 4.0 build 1005, 20/January/2003

- Splash screen image of the Standard ed. of BurnInTest corrected (indicated PRO).
- Number of logical processors (Hyperthreading) now displayed in log files.

Release 4.0 build 1004, 8/January/2004

NEW TESTS & IMPROVEMENTS TO EXISTING TESTS
A large number of the tests have been improved to be more comprehensive or
convenient to use. These are:
- CD-RW burn test added with Quick/Full format, 650/700MB and random
  seek options.
- USB 2.0 HighSpeed (480Mb/s) testing supported with PassMark's USB2Test plug.
- Looped Video playback test introduced, to detect codec support, playback 
  errors and dropped frames.
- New disk test modes including random data tests with random seeking and 
  butterfly seeking to more thoroughly exercise disk drives. The disk duty
  cycle may now be set per disk (with an override feature).
- New CD/DVD test modes including random data tests with random seeking.
- Maths/CPU test now supports test threads per CPU & CPU duty cycles scaling 
  more consistent with CPU utilisation. 
- MMX test now supports CPU duty cycles scaling more consistent with CPU 
  utilization. 
- Parallel Port test changed to allow automatic detection of non-sequential 
  LPT port numbers and non-standard IO addresses. In Windows 2000 and above, 
  automatic support of add-on Parallel ports (eg. using PCI cards) is also 
  provided. Retries have been added to the locking of the Parallel port. 
- Serial port test now includes baud rate cycling.
- Sound test now reports a greater range and detail of test error conditions.
- 2D test now reports a greater range and detail of test error conditions.
  A new 2D test option to skip 2D frame errors has been added.

IMPROVEMENTS TO TESTING FACILITIES
- The main BurnInTest errors are now categorised into, None, Information, 
  Warning, Serious and Critical. The error text and this classification may be 
  changed to better suit the customers test environment. Certain errors may 
  now be configured to be ignored as errors.
- Periodic disk Logging has been replaced with a more comprehensive and flexible 
  set of reporting/logging features. This includes:
  * User selectable reporting level (4 levels from summary to almost code debug).
  * Errors are now written in real time (as opposed to every x minutes).
  * Advanced log file naming now allows log files to be automatically prefixed 
    with user defined text or certain environment variables.
  * Additional information added to the reports, such as, a PASS/FAIL column per
    test and a summary of serious and critical errors per test or script run.
- Scripting is now more comprehensive. This includes:
  * Introduced new scripting commands, EXECUTE & EXECUTEWAIT, to allow external 
    applications to be started from a BurnInTest script.
  * Introduced new scripting commands, REBOOT & REBOOTEND, to allow reboots of 
    the PC to be included into scripts.
  * Introduced new scripting command, LOG, to generate a user text message in 
    the log history, to aid log interpretation.
  * Modification of scripting to display PASS/FAIL at the completion of the script.
  * Scripted testing may now be run from the command line.
  * Added a PASS/FAIL message to log files for script runs.
- Added automatic stop based on all tests reaching a user defined number of
  test cycles. Added a new scripting command, SETCYCLES, to also support this from
  a script.
- Support for additional 3rd party temperature monitoring software applications. 
  Intel Active Monitor, Mother Board Monitor (MBM) and SpeedFan.
- Extra command line options to position the main window and set the test duration.
- Improved the error reporting for disk and CD tests when there are insufficient 
  system resources. 
- A number of User interface improvements have been made: such as complete new look 
  main window, consistency of style across test windows, updated progress bars UI 
  colour changes, large stop button always on top, disk drive volume labels displayed, 
  ...

BUG CORRECTIONS
- Start time, test duration, stop time and Temperature results (eg. Minimum and 
  Maximum) now accumulated across multiple tests (when accumulate logs specified).
- Fixed a bug with network test preferences tab preventing very small values 
  being entered for the bad packet ratio.
- Correction of a bug where the failure to initialise DirectDraw would lead to 
  incorrect behaviour.
- AWE memory test progress bar changes for greater than 4GB memory.
- Corrected a bug that appeared when disk or CD tests were run with multiple disks, 
  followed by removing a disk in preferences, the previous results were mapped to 
  the incorrect drive.
- With a large number of tests, the bottom of the main window was overwritten on 
  the starting of tests. This has been corrected.
- Corrected a bug where the disk test may display a negative speed very briefly 
  at the start of the test.
- Corrected Preferences mouse over text.
- Corrected a bug where the USB port name was not correctly shown as "NA" when 
  the USB1 device drivers were not installed.
- Corrected a bug that prevented a scripted SLEEP after a floppy disk was tested.

History of earlier releases:
Please see http://passmark.com/products/bit_history.htm



Documentation
=============
All the documentation is included in the help file. It can be accessed from
the help menu. There is also a PDF format Users guide available for download 
from the PassMark web site.


Support
=======
For technical support, questions, suggestions, please check the help file for 
our email address or visit our web page at http://www.passmark.com


Ordering / Registration
=======================
All the details are in the help file documentation
or you can visit our sales information page
http://www.passmark.com/sales


Compatibility issues with the Network & Parallel Port Tests
===========================================================
If you are running Windows 2000 or XP, you need to have 
administrator privileges to run this test.


Enjoy..
The PassMark Development team
