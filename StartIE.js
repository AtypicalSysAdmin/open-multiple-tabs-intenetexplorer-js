var navOpenInBackgroundTab = 0x1000;
var objIE = new ActiveXObject("InternetExplorer.Application");

objIE.Navigate2("http://www.site1.com");
objIE.Navigate2("http://www.site2.com", navOpenInBackgroundTab);
objIE.Navigate2("http://www.site3.com", navOpenInBackgroundTab);
objIE.Visible = true;
 
