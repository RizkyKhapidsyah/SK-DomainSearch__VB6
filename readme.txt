DomainSearch - available from http://www.frez.co.uk

Should you wish to commission any consultancy work, please do not hesitate to contact us; sales@frez.co.uk

This code demonstrates how to search the INTERNIC and NOMINET registrars for available .COM and .CO.UK internet domains.

A number of files are stored in the application directory:

     CHECKED.TXT     - Stores a list of domains that have been checked so they
                       do not get checked again in the future
     PREFIX.TXT      - Text to add to the front of each domain name you want to
                       check so you can quickly check multiple domain names,
                       e.g. games.com, e-games.com, netgames.com
     SUFFIX.TXT      - Similar to PREFIX.TXT but addedd to the end of the domain
                       names
     AVAILABLE.TXT   - A list of the domains that have been searched that are
                       available
     UNAVAILABLE.TXT - A list of the domains that have been searched that are not
                       available
     UNKNOWN.TXT     - A list of the domains that have been searched whose status
                       could not be determined

This application now uses the OpenURL function provided by FreeVBCode.COM to work around a bug with the Microsoft Internet Transfer Control.

In addition, for each domain searched a file will be created in the format domain*.TXT that lists the domains that are available for that domain name.

This code could easily be extended to check other domains, for instance, INTERNIC is also responsible for the .ORG, .EDU and .NET domains, and NOMINET is also responsible for .ORG.UK, .NET.UK, .ORG.UK, .LTD.UK and the .PLC.UK domains. Other registrars could be added if required.

This is 'free' software with the following restrictions:

You may not redistribute this code as a 'sample' or 'demo'. However, you are free to use the source code in your own code, but you may not claim that you created the sample code. It is expressly forbidden to sell or profit from this source code other than by the knowledge gained or the enhanced value added by your own code.

Use of this software is also done so at your own risk. The code is supplied as is without warranty or guarantee of any kind.
