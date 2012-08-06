Transcription
=============

This is a tool to automate the copying of documents to an Excel spreadsheet, is written in Python with Tk.
You can set defaults ​and values ​​to fill (automcomplete fields and other options)




How to install
--------------

###### With virtualenv:
	$ python setup develop

###### Without virtualenv:
	# python setup develop




How to use
----------
###### Duplicate configuration template
	$ cp scripts/{transcription.cfg.default,transcription.cfg}
###### Edit configuration
	$ vim scripts/transcription.cfg
###### Run script
	$ python scrips/transcription_executable.py