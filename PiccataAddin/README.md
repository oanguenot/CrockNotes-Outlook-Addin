#Piccata Outlook Addin

This repository contains the source code of the Piccata Notebooks Outlook Addin.

This Outlook addin extract email information in an XML formant and add them to the clipboard for being pasted into Piccata Notebooks.

This Outlook addin can work without Piccata Notebooks. Data exported to the clipboard can be pasted into any others applications.

#Email XML format

The Email XML format is based on the following:

    <messageList>
       <message>
		   <fromDisplay>Olivier Anguenot</fromDisplay>
		   <fromAddress>olivier@piccatasoftware.com</fromAddress>
		   <subject>An E-mail from me</subject>
		   <deliveryTime>2012-08-10 12:37:43</deliveryTime>
		   <attachments>
			   <attachment>
				   <fileName>myFile.txt</fileName>
				   <fileType><fileType>
			   </attachment>
		   </attachments>
		   <contentText>
			   <![CDATA[
					Hello, this is an email from my mailbox
			   ]]>
		   </contentText>
	   </message>
    </messageList>
	
#Building

You must use Visual Studio 2008 at least to build the application

