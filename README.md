# Installation

## Mono

Install the latest version of Mono by following these instructions: [http://www.mono-project.com/download](http://www.mono-project.com/download). For an Ubuntu/Debian based distro the installation is as follows:

Add the Mono Project GPG signing key (if you don’t use sudo, be sure to switch to root):

`sudo apt-key adv --keyserver keyserver.ubuntu.com --recv-keys 3FA7E0328081BFF6A14DA29AA6A19B38D3D831EF`

Next, add the package repository (if you don’t use sudo, be sure to switch to root):

`echo "deb http://download.mono-project.com/repo/debian wheezy main" | sudo tee /etc/apt/sources.list.d/mono-xamarin.list`

Install Mono:

`sudo apt-get install mono-complete`

## Open XML SDK

Clone the Open XML SDK repository:

`git clone https://github.com/OfficeDev/Open-XML-SDK`

And `make` the library using the Makefile for Mono:

`make -f Makefile-Linux-Mono build`

## Compiling

Copy the Open XML library (`OpenXMLLib.dll` from the `build/` directory that was created in the previous step) into the current folder, and compile using `mcs`:

`mcs -r:OpenXMLLib.dll,WindowsBase.dll,System.Xml.Linq.dll RevisionAccepter.cs`

# Using

You can now run the binary on a Docx document like so:

`./RevisionAccepter document.docx`


