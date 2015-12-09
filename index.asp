<!--#include file="conn.asp" -->
<%Set rsQuery = Con.Execute("SELECT * FROM tblSkiTid") %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
	<title>Log ind</title>

</head>

<%if session("bemaerk") = "1" then %>
<body onload="javascript:alert('Din besked er sendt.')">
<%session("bemaerk") = "" %>
<%elseif session("pass-sendt") = "1" then %>
<body onload="javascript:alert('Såfremt der var oprettet en profil på den angivner mailadresse, \n er der sendt en mail med de(n) tilhørende password(s)')">
<%session("pass-sendt") = "" %>
<%else %>
<body>
<%end if %>
<div style="width:500px">
<img src="logo.gif" height="124" width="124" align="right">
<h1>Roskilde Skiklub</h1>
<h2>Tilmelding til Roskilde Festival</h2>
Festival starter om
<div id="test" style="color:red;font-weight:bold"></div>

<script type="text/javascript">
	if(!Date.prototype.getDaysInMonth) {
		// private
		Date.daysInMonth = [31,28,31,30,31,30,31,31,30,31,30,31];
	
		// public
		Date.prototype.getDaysInMonth = function() {
			Date.daysInMonth[1] = this.isLeapYear() ? 29 : 28;
			return Date.daysInMonth[this.getMonth()];
		};	
	}

	if(!Date.prototype.isLeapYear) {	
		// private
		Date.prototype.isLeapYear = function() {
			var year = this.getFullYear();
			return !!((year & 3) == 0 && (year % 100 || (year % 400 == 0 && year)));
		}
	}

	/**
	 * @class Countdown
	 * @param {string/HTMLElement} el target HTMLElement or id of target HTMLElement.
	 * @param {string/Date Object} time target time to countdown to.
	 */
	var CountDown = function(el, time) {
		if(!el.innerHTML) {
			el = document.getElementById(el);
		}	
		this.el = el;
		
		if(typeof time != "object" || !time.getTime) {
			time = new Date(time);
		}
		this.target = time;
	}
	
	CountDown.prototype = {
		dayText : 'dag',

		monthText : 'måned',
		
		yearText : 'år',
		
		hourText : 'time',
		
		minuteText : 'minut',
		
		secondText : 'sekund',
		
		plural1 : 'r',
		plural2 : 'er',
		plural3 : 'ter',	
		plural4 : 'e',		
		plural : '',
		/**
		 * Gets distance in seconds
		 * @return {number}
		 */
		getDist : function() {
			var now = new Date();
			
			var dist = now.getTime() - this.target.getTime();
			return Math.abs(dist);
		},

		/**
		 * Gets calculated distance
		 * @return {object}
		 */
		getCount : function() {
			var dist = this.getDist();
			
			var second = 1000
			var minute = 60 * second;
			var hour = 60 * minute;
			var day = 24 * hour;
			
			var days = Math.floor(dist / day);
			var hours = Math.floor((dist - (days*day)) / hour);
			var minutes = Math.floor((dist - (days*day+hours*hour)) / minute);
			var seconds = Math.floor((dist - (days*day+hours*hour+minutes*minute)) / second);
			
			var tmpTime = new Date();
			var tmpDays = 0, retractionDays = 0;
			var months = 0;
			
			while(days > tmpDays) {
				tmpTime.setMonth(tmpTime.getMonth()+1);
				tmpDays += tmpTime.getDaysInMonth();
				if(days < retractionDays + tmpTime.getDaysInMonth()) {
				} else {
					retractionDays += tmpTime.getDaysInMonth();
					months++;
				}
			}
			
			var years = Math.floor(months / 12);
			if(years > 0) {
				months = months - years * 12;
			}

			return {
			"years" : years, "months" : months, "days" : days-retractionDays, "hours" : hours, "minutes" : minutes, "seconds" : seconds
// hvis år skal vises
//"years" : years, "months" : months, "days" : days-retractionDays, "hours" : hours, "minutes" : minutes, "seconds" : seconds
			}
		},
		
		/**
		 * starts countdown
		 */
		start : function() {
			var c = this.getCount()
		
			var txt = [
				c.years + ' ' + this.yearText+(c.years != 1 ? this.plural : '')+ ' ',
				c.months + ' ' + this.monthText+(c.months != 1 ? this.plural2 : '')+ ' ',
				c.days + ' ' + this.dayText+(c.days != 1 ? this.plural4 : '')+ ' ',
				c.hours + ' ' + this.hourText+(c.hours != 1 ? this.plural1 : '')+ ' ',
				c.minutes + ' ' + this.minuteText+(c.minutes != 1 ? this.plural3 : '')+ ' ',
				c.seconds + ' ' + this.secondText+(c.seconds != 1 ? this.plural2 : '')+ ' '
			]


//			this.el.innerHTML =  txt.join('');
if (c.years > 0){
			this.el.innerHTML =  c.years + ' ' + this.yearText+(c.years != 1 ? this.plural : '')+ ', ' + c.months + ' ' + this.monthText+(c.months != 1 ? this.plural2 : '')+ ', ' +
				c.days + ' ' + this.dayText+(c.days != 1 ? this.plural4 : '')+ ', ' +
				c.hours + ' ' + this.hourText+(c.hours != 1 ? this.plural1 : '')+ ', ' +
				c.minutes + ' ' + this.minuteText+(c.minutes != 1 ? this.plural3 : '')+ ' og ' +
				c.seconds + ' ' + this.secondText+(c.seconds != 1 ? this.plural2 : '')+ ' '
}

if (c.years < 1){
			this.el.innerHTML =  c.months + ' ' + this.monthText+(c.months != 1 ? this.plural2 : '')+ ', ' +
				c.days + ' ' + this.dayText+(c.days != 1 ? this.plural4 : '')+ ', ' +
				c.hours + ' ' + this.hourText+(c.hours != 1 ? this.plural1 : '')+ ', ' +
				c.minutes + ' ' + this.minuteText+(c.minutes != 1 ? this.plural3 : '')+ ' og ' +
				c.seconds + ' ' + this.secondText+(c.seconds != 1 ? this.plural2 : '')+ ' '
}

if (c.months < 1){
			this.el.innerHTML =  c.days + ' ' + this.dayText+(c.days != 1 ? this.plural4 : '')+ ', ' +
				c.hours + ' ' + this.hourText+(c.hours != 1 ? this.plural1 : '')+ ', ' +
				c.minutes + ' ' + this.minuteText+(c.minutes != 1 ? this.plural3 : '')+ ' og ' +
				c.seconds + ' ' + this.secondText+(c.seconds != 1 ? this.plural2 : '')+ ' '
}


			
			var fn = this._delegate(this.start,this);
			this.timer = setTimeout(fn,1000);
		},

		/**
		 * stops countdown
		 */
		stop : function() {
			clearTimeout(this.timer);
		},
		
		// private
		_delegate : function(func,obj) {
			return function() {
				return func.apply(obj,[]);
			}
		}
	}

	window.onload = function() {
		// Start the countdown format mm/dd/yyyy hh:mm:ss
		var cd = new CountDown("test", new Date("07/01/2010 16:00:00"));
		cd.start();
		
	}
</script>
<br>
</div>



<%if rsQuery("tid") = "1" then%>

Du har følgende valgmuligheder:
<hr>
<strong><em>1. Opret en tilmelding</em></strong><br>
Her kan du <a href="opret.asp">oprette en tilmelding</a> og dermed ønske vagter.
<br>
Vi skal bruge i alt 225 medarbejdere. Når vi har fået det ønskede antal tilmeldinger - vil der blive oprettet en venteliste.
<br>
Bemærk, at du skal være minimum 15 år for at kunne tilmelde dig.
<hr>
<strong><em>2. Redigere dit valg af vagter</em></strong><br>
Hvis du har rettelser til allerede ønskede vagter, kan du indtil søndag den 23. maj 2010 logge på her. 
<br>
<form id="logon" action="login.asp" method="get">
Email: <input type="text" name="email"><br>
Password: <input type="password" name="pw"><br>
<input type="submit" value="Log på">
</form>
<hr>
<strong><em>3. Få tilsendt dit password</em></strong><br>
Hvis du har glemt dit password, kan du indskrive din emailadresse herunder.<br>
Såfremt adressen er tilknyttet en eller flere profiler, sender vi de(t) tilknyttede password(s)
<br>
<form id="sendPass" action="sendPass.asp" method="get">
Email: <input type="text" name="email"><br>
<input type="submit" value="Send">
</form>
<hr>

<%end if %>

<%if rsQuery("tid") = "2" then%>
<br>

Der er kortvarigt lukket for tilmelding pga. administrative opdateringer - prøv venligst igen lidt senere!
<br>
<br>
<hr>

<%end if %>





<%if rsQuery("tid") = "3" then%>
Det er ikke længere muligt at tilmelde sig vagter ved Roskilde Festivalen, eller redigere de vagtønsker du har angivet.<br>
Du kan dog stadig:
<hr>
<strong><em>1. Se hvilke vagter du har tilmeldt dig.</em></strong><br>
Log på med din email og dit password.
<form id="logon" action="login.asp" method="get">
Email: <input type="text" name="email"><br>
Password: <input type="password" name="pw"><br>
<input type="submit" value="Log på">
</form>
<hr>
<strong><em>2. Få tilsendt dit password</em></strong><br>
Hvis du har glemt dit password, kan du indskrive din emailadresse herunder.<br>
Såfremt adressen er tilknyttet en eller flere profiler, sender vi de(t) tilknyttede password(s)
<br>
<form id="sendPass" action="sendPass.asp" method="get">
Email: <input type="text" name="email"><br>
<input type="submit" value="Send">
</form>
<hr>

<%end if %>


<%if rsQuery("tid") = "4" then%>
<strong><em>1. Se de vagter du har fået tildelt.</em></strong><br>
Log på med din email og dit password.
<form id="logon" action="login.asp" method="get">
Email: <input type="text" name="email"><br>
Password: <input type="password" name="pw"><br>
<input type="hidden"  name="tid" value="3">
<input type="submit" value="Log på">
</form>
<hr>
<strong><em>2. Få tilsendt dit password</em></strong><br>
Hvis du har glemt dit password, kan du indskrive din emailadresse herunder.<br>
Såfremt adressen er tilknyttet en eller flere profiler, sender vi de(t) tilknyttede password(s)
<br>
<form id="sendPass" action="sendPass.asp" method="get">
Email: <input type="text" name="email"><br>
<input type="submit" value="Send">
</form>
<hr>
<%end if %>


<%if rsQuery("tid") = "5" then%>

Det er ikke muligt at indtaste eller se noget på denne side i øjeblikket. Siden vil være åben igen fra starten af april og indtil Roskilde Festival.
<br>
<br>
<hr>


<%end if %>

</body>
</html>
