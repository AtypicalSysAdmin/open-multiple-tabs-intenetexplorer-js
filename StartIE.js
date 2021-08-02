var navOpenInBackgroundTab = 0x1000;
var objIE = new ActiveXObject("InternetExplorer.Application");

objIE.Navigate2("http://www.google.com");
objIE.Navigate2("http://www.microsoft.com", navOpenInBackgroundTab);
objIE.Navigate2("http://www.amazon.com", navOpenInBackgroundTab);
objIE.Visible = true;
 
