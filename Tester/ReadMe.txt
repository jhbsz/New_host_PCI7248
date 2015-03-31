
==== Alcor Micro Smart Card Reader Demo Program for EEPROM card operation

This is a demo program which provides a way to operate the eeprom card using Alcor's smartcard reader.


==== Using SmartCard APIs
The complete procedure to operate the EEPROM card by SmartCard APIs :
SCardEstablishContext (.....) ==> Connect to the smart card resource manager
	SCardConnectA(.....) ==> Connect to the smart card reader
		SCardControl(..,IOCTL_SET_UNRESPONSED_CARD_TO_MCARD,...) ==> tell the alcor driver to accept an EEPROM card
			SCardBeginTransaction(....)==> block other applications access the card
				SCardControl(..,IOCTL_I2C_COMMAND,..) ==> switch the reader to I2C mode
				SCardControl(..,IOCTL_I2C_READ,..) ==> Read data from EEPROM card
				SCardControl(..,IOCTL_I2C_WRITE,..) ==> Write data to EEPROM card
			SCardEndTransaction(....) ==> undo SCardBeginTransaction(....)
		SCardControl(..,IOCTL_CLEAR_UNRESPONSED_CARD_TO_MCARD,...) ==>undo SCardControl(..,IOCTL_SET_UNRESPONSED_CARD_TO_MCARD,...)
	SCardDisconnectA(.....) ==> Disconnect the smartcard reader
SCardReleaseContext(.....) ==> Release the connection with the resource manger

Also, we use a timer control to polling the reader state :
SCardGetStatusChange(...) ==> check if the state was changed from the resource manager
SCardStatusA(..) ==> we don't use this API in the version

Useing SCardGetStatusChange(...) in a periodly polling routine is better than SCardStatusA(..),
because it only request the reader state from the resource, not from the driver, that will
save the system resource very much.



 

