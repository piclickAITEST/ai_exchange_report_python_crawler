#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#include COM_L.ahk
#include IE_Macro_By_Tag_L.ahk
GUI,Color,White

Gui, Add, GroupBox, x12 y30 w470 h390 , 

;Gui, Add, Text, x22 y40 w410 h60 vtargetURL, 아래 각 입력창은 각 게시판의 긁어올 가장 뒷 페이지 값을 입력합니다. 3을 입력하면 3페이지부터 가장 최신 글까지 긁어 옵니다.

Gui, Add, Text, x22 y100 w110 h20 vtText1,  지마켓 크롤러
Gui, Add, Button, x222 y95 w80 h20 g창띄우기, 창띄우기
;Gui, Add, Button, x222 y125 w80 h20 g정보수집, 정보수집


; Generated using SmartGUI Creator 4.0
Gui, Show, w500 h520, 지마켓 크롤러..



Enabled := ComObjError(false)


return



saveSetting:
	
	Gui,submit,nohide

	IniWrite, %ID%,  dripzil.ini,  로그인,ID
	IniWrite, %PW%,  dripzil.ini,  로그인,PW
	IniWrite, %일반%,  dripzil.ini,  설정,일반
	IniWrite, %HTML%,  dripzil.ini,  설정,HTML
	IniWrite, %delay%,  dripzil.ini,  설정,delay
	IniWrite, %모두%,  dripzil.ini,  설정,모두
	IniWrite, %SerialKey%,  dripzil.ini,  설정,SerialKey
	IniWrite, %CrawlURL%,  dripzil.ini,  설정,CrawlURL

	return

RegRead(RootKey, SubKey, ValueName = "") {
   RegRead, v, %RootKey%, %SubKey%, %ValueName%
   Return, v
}

isloadComplete(driver)
{
	result := driver.executeScript("return (document.readyState == 'loaded' || document.readyState == 'complete');")
	return result
}

CreateInsertScriptAndDownloadTorrent(driver, CrawlURL) {
		
	;msgbox, create script

	fileName := "kukudasyumoContent.info"

	;msgbox, % fileName

	FileDelete, %fileName%
	
	;fileName := tmpFn ".list"

	i := 0
	waitTime := 1

	titleList := Object()
	linkList := Object()
	imageList := Object()
	iframesrcList := Object()
	descList := Object()

	dataList := Object()
	;msgbox, hhh
	Loop, Read, kukudasyumoContent.list ; This loop retrieves each line from the file, one at a time.
	{
		linkList[i] := A_LoopReadLine
		i++
	}

	contentsCnt := i

	;msgbox, % contentsCnt

	; step 2 : get real iframe src....
	; 필요하면, 파일에서 읽어 오자...
	
	;for i, linkURL in linkList
	i := contentsCnt - 1
	;fileName := tmpFn ".info"
	;playbuttonimg := "<img src=""http://www.getlucky.co.kr/img/playnow.jpg"" width=""100%"" height=""auto"">"

	loop
	{
		if ( i < 0 )
			break
		
		linkURL := linkList[i]

		driver.Get(linkURL)

		k := 0
		loop {
			ret := driver.executeScript("return document.readyState;")
			;if (ret = "loaded" or ret = "complete")
			if (ret = "complete")
				break

			;msgbox, % ret
			if( k > 14 ) ; 최대 7초만 기다려 본다.
				break
			sleep, 500
			k++
		}
		sleep, 1000

		gameDesc := driver.executeScript("return document.querySelector('div.view-wrap h1').innerText;")
		gameDesc := RegExReplace(gameDesc,"'","\'")
		gameDesc := RegExReplace(gameDesc,"`n","")
		;descList[i] := gameDesc

		;msgbox, % gameDesc

		embedsrc := driver.executeScript("return document.querySelector('div.view-content').outerHTML;") ; pwb.document.querySelector("
		embedsrc := RegExReplace(embedsrc,"'","\'")
		embedsrc := RegExReplace(embedsrc,"`n","")

		imgsrc := driver.executeScript("return document.querySelector('div.view-content img').src;") ; pwb.document.querySelector("div.view-content img").src
		imgsrc := RegExReplace(imgsrc,"'","\'")

		;msgbox, % imgsrc
		imgFileName := "yumo_img\jp_yumo_" i ".jpg"

		;msgbox, % imgFileName
		;URLDownloadToFile,%imgsrc%,%imgFileName%

		contentstext := driver.executeScript("return document.querySelector('div.view-content').innerText;")
		contentstext := RegExReplace(contentstext,"'","\'")

		;msgbox, % embedsrc

		j := 0
		torrentFileName := ""
		torrentURL := ""
		linkCnt := driver.executeScript("return document.querySelectorAll('div.list-group a').length;")

		if ( linkCnt = 0 ) {
			i--
			continue
		}

		loop {
			if (j = linkCnt)
				break
			
			torrentFileName := driver.executeScript("return document.querySelectorAll('div.list-group a')[" j "].innerText;")
			torrentURL := driver.executeScript("return document.querySelectorAll('div.list-group a')[" j "].href;")

			IfInString, torrentFileName, .torrent
				break
			
			j++
		}
		tmpStr := driver.executeScript("return document.querySelectorAll('div.list-group a span')[0].innerText;")
		StringReplace,torrentFileName,torrentFileName,%tmpStr%,

		tmpStr := driver.executeScript("return document.querySelectorAll('div.list-group a span')[1].innerText;")
		StringReplace,torrentFileName,torrentFileName,%tmpStr%,

		torrentFileName := trim(RegExReplace(torrentFileName,".torrent(.*)",".torrent"))

		torrentFileName := "yumo_torrent\" torrentFileName

		iframesrcList[i] := embedsrc

		; download torrent...
		driver.get(torrentURL)
		sleep, 1000


		;COM_Invoke(pwb, "Navigate", "document.querySelectorAll('div.list-group a')[j].click();" , 0x0400)

		;URLDownloadToFile,%torrentURL%,%torrentFileName%
	
		;msgbox, % torrentFileName
		
		strLine := gameDesc "`t" contentstext "`t" imgFileName "`t" torrentFileName "`t" embedsrc

		;strLine :=  "insert into g5_write_quizNvote (wr_subject, wr_content) values ('" titleList[i] "', '" contentsData "');"
		
		;msgbox, % strLine

		FileAppend, %strLine% `n, %fileName% 

		i--
	}

}

waitBrowser(driver) {
	k := 0
	loop {
		ret := driver.executeScript("return document.readyState;")
		;if (ret = "loaded" or ret = "complete")
		if (ret = "complete")
			break

		;msgbox, % ret
		if( k > 14 ) ; 최대 7초만 기다려 본다.
			break
		sleep, 500
		k++
	}
}

CheckloadComplete(driver)
{
	k := 0
	loop {
		ret := driver.executeScript("return document.readyState;")
		if (ret = "complete")
			break

		if( k > 20 ) ; 최대 10초만 기다려 본다.
			break
		sleep, 500
		k++
	}
}


WriteBoard(pwb, title, body, boardURL) {

	pwb.Navigate(boardURL)
	IE_Loading_Check(pwb) 
	sleep, 1000
	;msgbox, load board url

	COM_Invoke(pwb, "Navigate", "javascript:document.getElementById('html').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,1000

	pwb.document.getElementById("wr_subject").value := title
	;COM_Invoke(pwb, "Navigate","javascript:document.querySelector('button.se2_to_html').click();", 0x0400)
	;IE_Loading_Check(pwb) 
	;sleep,1000
	;msgbox, html click
	
	;pwb.document.querySelector("textarea.se2_input_syntax.se2_input_htmlsrc").value := body
	pwb.document.getElementById("wr_content").value := body
	COM_Invoke(pwb, "Navigate", "javascript:document.getElementById('btn_submit').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,1000

}

WriteBoardBasic(pwb, title, body, boardURL) {

	pwb.Navigate(boardURL)
	IE_Loading_Check(pwb) 
	sleep, 2000
	;msgbox, load board url
	
	;html 편집 클릭
	COM_Invoke(pwb, "Navigate", "javascript:document.querySelector('div.form-group iframe').contentWindow.document.querySelector('button.se2_to_html').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,2000

	pwb.document.getElementById("wr_subject").value := title
	;COM_Invoke(pwb, "Navigate","javascript:document.querySelector('button.se2_to_html').click();", 0x0400)
	;IE_Loading_Check(pwb) 
	;sleep,1000
	;msgbox, html click
	
	;pwb.document.querySelector("textarea.se2_input_syntax.se2_input_htmlsrc").value := body
	pwb.document.querySelector("div.form-group iframe").contentWindow.document.querySelector("textarea.se2_input_syntax.se2_input_htmlsrc").value := body
	;pwb.document.getElementById("wr_content").value := body
	COM_Invoke(pwb, "Navigate", "javascript:document.getElementById('btn_submit').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,1000

}
fl
WriteContentsVideo(pwb, driver, page, boardURL, baseURL) {
		
	;msgbox, load crawl 

	;page := 5
	prepage := page-1	
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman01&page="
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman02&page="

	; step 1 : 현재 페이지의 가장 최신 글 제목을 가져온다....

	boardlisturl := RegExReplace(boardURL,"write.php","board.php")
	pwb.Navigate(boardlisturl)
	IE_Loading_Check(pwb) 
	sleep, 1000
	;msgbox, load board url

	rowcnt := pwb.document.querySelector(".list-body").querySelectorAll(".list-item").length
	latestsubject := ""

	if(rowcnt > 0) {
		basesubject := pwb.document.querySelectorAll("div.wr-subject")[0].innerText	
	}

	if(latestsubject) {
		msgbox, % basesubject

	}


	pwb.Navigate(boardURL)
	IE_Loading_Check(pwb) 
	sleep, 1000
	
	;html 편집 클릭
	COM_Invoke(pwb, "Navigate", "javascript:document.querySelector('div.form-group iframe').contentWindow.document.querySelector('button.se2_to_html').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,1000

	driver.get(baseURL prepage "&page=" page)
	CheckloadComplete(driver)
	sleep, 1000

	Loop
	{
		if( page < 1 )
			break

		listPage := baseURL page
		; 목록 페이지로 이동...
		driver.get(listPage)
		CheckloadComplete(driver)
		sleep, 1000

		;document.querySelectorAll('table.bd_lst > tbody > tr')[0].querySelector('td.title a').click();

		
		i := driver.executeScript("return document.querySelectorAll('tr > td.listnum + td + td + td.listnum + td.listnum + td.listnum').length;")
		i--
		
		;msgbox, % i

		Loop
		{
			if( i < 0 )
				break
			
			;msgbox, move to detail view...

			driver.executeScript("document.querySelectorAll('tr > td.listnum + td + td + td.listnum + td.listnum + td.listnum')[" i "].parentElement.querySelector('td + td > a').click();")
			CheckloadComplete(driver)
			sleep, 1000
			

			
			title := driver.executeScript("return document.querySelector('td > table + table').querySelector('tbody > tr + tr + tr + tr + tr + tr ').querySelector('td + td').innerText;")

			if ( title = "" )
				title := "무제"

			;msgbox, % title

			sleep, 500

			
			body := driver.executeScript("return document.querySelector('td > table + table + table > tbody > tr > td').innerHTML;")

			;msgbox, % body

			WriteBoardBasic(pwb, title, body, boardURL)

			i--

			; 다시 목록 페이지로 간다.
			driver.get(listPage)
			CheckloadComplete(driver)
			sleep, 1000
		}

		page--
		prepage--
	}
}

WriteContentsnsbu(pwb, pwb2, driver, page, boardURL, boardURL2, baseURL) {
		
	;msgbox, load crawl 

	;page := 5
	prepage := page-1	
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman01&page="
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman02&page="

	; step 1 : 현재 페이지의 가장 최신 글 제목을 가져온다....

	boardlisturl := RegExReplace(boardURL,"write.php","board.php")
	pwb.Navigate(boardlisturl)
	IE_Loading_Check(pwb) 
	sleep, 1000
	;msgbox, load board url

	rowcnt := pwb.document.querySelector(".list-body").querySelectorAll(".list-item").length
	latestsubject := ""

	if(rowcnt > 0) {
		basesubject := pwb.document.querySelectorAll("div.wr-subject")[0].innerText	
	}

	if(latestsubject) {
		msgbox, % basesubject

	}


	pwb.Navigate(boardURL)
	IE_Loading_Check(pwb) 
	sleep, 1000
	
	;html 편집 클릭
	COM_Invoke(pwb, "Navigate", "javascript:document.querySelector('div.form-group iframe').contentWindow.document.querySelector('button.se2_to_html').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,1000

	driver.get(baseURL prepage "&page=" page)
	CheckloadComplete(driver)
	sleep, 1000

	Loop
	{
		if( page < 1 )
			break

		listPage := baseURL page
		; 목록 페이지로 이동...
		driver.get(listPage)
		CheckloadComplete(driver)
		sleep, 1000

		i := driver.executeScript("return document.querySelectorAll('tbody > tr').length;")
		i--
		
		Loop
		{
			if( i < 0 )
				break
			
			;msgbox, move to detail view...

			driver.executeScript("document.querySelectorAll('tbody > tr')[" i "].querySelector('td.list-subject a').click();")
			CheckloadComplete(driver)
			sleep, 1000
			

			; remove good button...
			;driver.executeScript("document.querySelector('div.print-hide.view-good-box').remove();")
			;sleep, 500
			;driver.executeScript("document.querySelector('h2.board-view-atc-title').remove();")
			;sleep, 500		  document.querySelector('div.top_area').innerText
			
			title := driver.executeScript("return document.querySelector('.view-subject').innerText;")

			if ( title = "" ) {
				title := driver.executeScript("return document.querySelector('div.border_title > h1').innerText;")
				if(title = "")
					msgbox, no title...
			}
			;msgbox, % title

			sleep, 500
			
			driver.executeScript("$('img[src*=namsungbu]').each(function() {this.remove();});")
			driver.executeScript("$('img[src*=nsbu]').each(function() {this.remove();});")
			driver.executeScript("$('img[src*=manpeace]').each(function() {this.remove();});")

			driver.executeScript("$('img[src*=editor]').each(function() {this.remove();});")

			body := driver.executeScript("return document.querySelector('div.view-content').innerHTML;")
			;msgbox, % body

			WriteBoardBasic(pwb, title, body, boardURL)
			if(boardURL2)
				WriteBoardBasic(pwb2, title, body, boardURL2)

			i--

			; 다시 목록 페이지로 간다.
			driver.get(listPage)
			CheckloadComplete(driver)
			sleep, 1000
		}

		page--
		prepage--
	}
}

WriteContentsDasi(pwb,pwb2, driver, page, boardURL, boardURL2, baseURL) {
		
	;msgbox, load crawl 

	;page := 5
	prepage := page-1	
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman01&page="
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman02&page="

	; step 1 : 현재 페이지의 가장 최신 글 제목을 가져온다....

	boardlisturl := RegExReplace(boardURL,"write.php","board.php")
	pwb.Navigate(boardlisturl)
	IE_Loading_Check(pwb) 
	sleep, 1000
	;msgbox, load board url

	rowcnt := pwb.document.querySelector(".list-body").querySelectorAll(".list-item").length
	latestsubject := ""

	if(rowcnt > 0) {
		basesubject := pwb.document.querySelectorAll("div.wr-subject")[0].innerText	
	}

	if(latestsubject) {
		msgbox, % basesubject

	}


	pwb.Navigate(boardURL)
	IE_Loading_Check(pwb) 
	sleep, 1000
	
	;html 편집 클릭
	COM_Invoke(pwb, "Navigate", "javascript:document.querySelector('div.form-group iframe').contentWindow.document.querySelector('button.se2_to_html').click();", 0x0400)
	IE_Loading_Check(pwb) 
	sleep,1000

	;pwb.document.getElementById("wr_subject").value := title




	

	driver.get(baseURL prepage "&page=" page)
	CheckloadComplete(driver)
	sleep, 1000

	Loop
	{
		if( page < 1 )
			break

		listPage := baseURL page
		; 목록 페이지로 이동...
		driver.get(listPage)
		CheckloadComplete(driver)
		sleep, 1000

		;document.querySelectorAll('table.bd_lst > tbody > tr')[0].querySelector('td.title a').click();

		
		i := driver.executeScript("return document.querySelectorAll('table.bd_lst > tbody > tr').length;")
		i--
		
		;msgbox, % i

		Loop
		{
			if( i < 0 )
				break
			
			driver.executeScript("document.querySelectorAll('table.bd_lst > tbody > tr')[" i "].querySelector('td.title a').click();")
			CheckloadComplete(driver)
			sleep, 1000
			;msgbox, move to detail view...

			; remove good button...
			;driver.executeScript("document.querySelector('div.print-hide.view-good-box').remove();")
			;sleep, 500
			;driver.executeScript("document.querySelector('h2.board-view-atc-title').remove();")
			;sleep, 500		  document.querySelector('div.top_area').innerText
			
			title := driver.executeScript("return document.querySelector('div.top_area').innerText;")

			if ( title = "" )
				title := "무제"

			;msgbox, % title

			sleep, 500

			driver.executeScript("document.querySelector('div.rd_body').querySelector('div.ad728').remove();")
			
			body := driver.executeScript("return document.querySelector('div.rd_body').outerHTML;")

			;msgbox, % body

			WriteBoardBasic(pwb, title, body, boardURL)
			if(boardURL2)
				WriteBoardBasic(pwb2, title, body, boardURL2)

			i--

			; 다시 목록 페이지로 간다.
			driver.get(listPage)
			CheckloadComplete(driver)
			sleep, 1000
		}

		page--
		prepage--
	}
}

WriteContents(pwb, pwb2, driver, page, boardURL, boardURL2, baseURL) {
		
	;msgbox, load crawl 

	;page := 5
	prepage := page-1	
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman01&page="
	;baseURL := "http://dripzil.com/bbs/board.php?bo_table=woman02&page="
	driver.get(baseURL prepage "&page=" page)
	CheckloadComplete(driver)
	sleep, 1000

	Loop
	{
		if( page < 1 )
			break

		listPage := baseURL page
		; 목록 페이지로 이동...
		driver.get(listPage)
		CheckloadComplete(driver)
		sleep, 1000

		
		i := driver.executeScript("return document.querySelectorAll('td.list-subject a').length;")
		i--

		;msgbox, % i

		Loop
		{
			if( i < 0 )
				break
			
			driver.executeScript("document.querySelectorAll('td.list-subject a')[" i "].click();")
			CheckloadComplete(driver)
			sleep, 1000
			;msgbox, move to detail view...

			; remove good button...
			;driver.executeScript("document.querySelector('div.print-hide.view-good-box').remove();")
			;sleep, 500
			;driver.executeScript("document.querySelector('h2.board-view-atc-title').remove();")
			;sleep, 500		
			
			title := driver.executeScript("return document.querySelector('div.border_title h1').innerText;")

			if ( title = "" )
				title := "무제"

			sleep, 500
			
			body := driver.executeScript("return document.querySelector('div.view-content').outerHTML;")

			;msgbox, % body

			WriteBoardBasic(pwb, title, body, boardURL)
			if(boardURL2)
				WriteBoardBasic(pwb2, title, body, boardURL2)


			i--

			; 다시 목록 페이지로 간다.
			driver.get(listPage)
			CheckloadComplete(driver)
			sleep, 2000
		}

		page--
		prepage--
	}
}
DoLoginChrome(driver, url, mallid, userid, passwd) {
	
	driver.get(url)
	CheckloadComplete(driver)
	sleep, 2000

	;return

	login_script =
	(
		document.getElementsByName('mall_id')[0].value="%mallid%"; 
		document.getElementsByName('userid')[0].value="%userid%"; 
		document.getElementsByName('userpasswd')[0].value = "%passwd%";
		;document.querySelector('div.form-body form button').click();
	) 
	driver.executeScript(login_script)
	CheckloadComplete(driver)
	sleep, 1000
}

DoLogin(pwb, url, userid, passwd) {
	
	pwb.Navigate(url)
	IE_Loading_Check(pwb) 
	sleep,3000		
	; anonymous / kiss2me!
	login_script =
	(
		javascript:
		document.getElementsByName('mall_id')[0].value='benitomaster';
		document.getElementsByName('userid')[0].value='%userid%';
		document.getElementsByName('userpasswd')[0].value='%passwd%' ;
		;document.querySelector('div.form-body form button').click();
	) 

	COM_Invoke(pwb, "Navigate", login_script , 0x0400)
	IE_Loading_Check(pwb) 
	sleep,3000		
	
	return
}


/*
<p> Help the Pirate in this Klondike game. Move all cards to the top 4 foundations. On the tableau built down in alternating color. Click on the top left stack to get a new open card. 
<p><p>
<img src="https://www.htmlgames.com/uploaded/thumb/piratesolitaire300.jpg" width="100%" height=auto>
<p><p>
<img src="http://www.getlucky.co.kr/img/playnow.jpg" width="100%" height=auto>
*/
창띄우기:
	Gui,submit,nohide

	;Step 1 : 라이센스 체크 및 다음 로그인하기...
	if(!driver) {
		;msgbox "create object"
		driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
		driver.AddArgument("--window-size=1800,1000")
		DoLoginChrome(driver, "https://eclogin.cafe24.com/Shop/?url=Init&login_mode=2&is_multi=F", "benitomaster", "guest", "qpslxh88!!")	
	}

	if(!driver2) {
		;msgbox "create object"
		driver2:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
		driver2.AddArgument("--window-size=1500,768")
		;DoLoginChrome(driver, "https://eclogin.cafe24.com/Shop/?url=Init&login_mode=2&is_multi=F", "benitomaster", "guest", "qpslxh88!!")	
	}
	
	msgbox "click when login ready..."

	getOrderInformation(driver, driver2, 98)
	return

getOrderInformation(driver, driver2, site_id) {
	Gui,submit,nohide
	
	year :=  % A_YYYY
	month :=   % A_MM
	day :=  % A_DD

	refdate := % year . "-" . month . "-" . day
	
	idx := 0
	loop {
		day--
		if(day < 1) {
			day := day + 31
			month--
			if(month = 9) {
				day := 30
			}
			if(month = 11) {
				day := 30
			}
		}
		if(month < 1) {
			month := month + 12
			year--
		}

		daystr := day
		monthstr := month
		if(day < 10) { 
			daystr := "0" day 
		}
		if(month < 10) { 
			monthstr := "0"month 
		}
		
		refdate := year "-" monthstr "-" daystr

		;msgbox % refdate
		
		idx++

		if(idx > 30) {
			break
		}

		; 전체 주문수
		url := "https://benitomaster.cafe24.com/admin/php/shop1/s_new/order_list_item.php?rows=20&searchSorting=order_desc&isBusanCall=&isChinaCall=&orderCallnum=&cticall=&realclick=T&tabclick=F&MSK[]=order_id&MSV[]=&orderStatusPayment=all&date_type=order_date&btnDate=7&product_search_type=product_name&find_option=product_no&order_product_name=&order_product_code=&order_product_no=&order_product_text=&order_set_product_no=&layer_order_product_code=&layer_order_product_opt_id=&popup_item_code=&popup_product_code=&payed=&payed_sql_version=&bank_info=&memberType=1&group_no=&isMemAuth=&isBlackList=&isFirstOrder=&isPointfyUsedMember=&shipment_type=all&bunch=&shippedAgain=&shipmentMessage=&delivSeperated=&isReservedOrder=&isSubscriptionOrder=&paystandard=choice&product_total_price1=&product_total_price2=&item_count_start=&item_count_end=&orderPathType=A&search_SaleOpenMarket[]=cafe24&search_SaleOpenMarket[]=mobile&search_SaleOpenMarket[]=mobile_d&search_SaleOpenMarket[]=NCHECKOUT&search_SaleOpenMarket[]=gmarket&search_SaleOpenMarket[]=auction&search_SaleOpenMarket[]=sk11st&search_SaleOpenMarket[]=shopn&search_SaleOpenMarket[]=inpark&search_SaleOpenMarket[]=coupang&search_SaleOpenMarket[]=kakao&search_SaleOpenMarket[]=womanstalk&search_SaleOpenMarket[]=tenten&search_SaleOpenMarket[]=wemake&search_SaleOpenMarket[]=melchi&search_SaleOpenMarket[]=halfclub&search_SaleOpenMarket[]=boribori&search_SaleOpenMarket[]=ogage&search_SaleOpenMarket[]=moongori&search_SaleOpenMarket[]=shopeesg&search_SaleOpenMarket[]=shopeeid&search_SaleOpenMarket[]=shopeemy&search_SaleOpenMarket[]=shopeetw&search_SaleOpenMarket[]=shopeeth&search_SaleOpenMarket[]=shopeeph&search_SaleOpenMarket[]=brich&search_SaleOpenMarket[]=zigzag&search_SaleOpenMarket[]=ably&search_SaleOpenMarket[]=timon&search_SaleOpenMarket[]=musinsa&search_SaleOpenMarket[]=wizwid&search_SaleOpenMarket[]=hottracks&search_SaleOpenMarket[]=akmall&search_SaleOpenMarket[]=daisomall&search_SaleOpenMarket[]=lfmall&search_SaleOpenMarket[]=styleshare&search_SaleOpenMarket[]=aland&search_SaleOpenMarket[]=rakutenkr&search_SaleOpenMarket[]=cjmall&search_SaleOpenMarket[]=lotteon&search_SaleOpenMarket[]=himart&search_SaleOpenMarket[]=tofkof&search_SaleOpenMarket[]=MORUGI&search_SaleOpenMarket[]=11st&mkSaleType=M&mkSaleTypeChg=&inflowPathType=A&inflowPathDetail=0000000000000000000000000000000000&paymethodType=A&paymentMethod[]=cash&paymentMethod[]=card&paymentMethod[]=tcash&paymentMethod[]=icash&paymentMethod[]=cell&paymentMethod[]=deferpay&paymentMethod[]=cvs&paymentMethod[]=point&paymentMethod[]=mileage&paymentMethod[]=deposit&paymentMethod[]=etc&pgListType=A&pgList[]=danal&pgList[]=dacom&pgList[]=payco&pgList[]=paynow&pgList[]=smilepay&pgList[]=eximbay&pgList[]=etc&paymentInfo=&discountMethod=&shop_no_order=1&delvReady=&delvCancel=&orderStatusNotPayCancel=N&orderStatusCancel=N&orderSearchCancelStatus=&orderStatusExchange=N&orderSearchExchangeStatus=&orderStatusReturn=N&orderStatusRefund=N&orderSearchRefundStatus=&orderSearchShipStatus=&orderStatus[]=all&orderStatus[]=N10&orderStatus[]=N20&orderStatus[]=N22&orderStatus[]=N21&orderStatus[]=N30&orderStatus[]=N40&RefundType=&RefundSubType=&sc_id=&second_shipping_company_id=&HopeShipCompanyId=all&post_express_flag=&tabStatus=&paymethod_total_count=&search_invoice_print_flag=all&search_is_escrow_shipping_registered=all&search_print_second_invoice=all&incoming=&is_purchased=&order_fail_code=&isBlackOrder=&start_date=" refdate "&year1=" year "&month1=" monthstr "&day1=" daystr "&start_time=00:00:00&end_date=" refdate "&year2=" year "&month2=" monthstr "&day2=" daystr "&end_time=23:59:59&realclick=T"
		driver.get(url)
		sleep,2000
		totalcnt := driver.executeScript("return document.querySelector('p.total strong').innerText")
		;msgbox % totalcnt	
		; save to db...
		url := "https://log.piclick.kr/util/exchange_stat.php?type=order&cnt=" totalcnt "&date=" refdate "&site_id=" site_id
		driver2.get(url)
		;msgbox % url

		; 전체 반품수
		url := "https://benitomaster.cafe24.com/admin/php/shop1/s_new/order_list_item.php?rows=20&searchSorting=order_desc&isBusanCall=&isChinaCall=&orderCallnum=&cticall=&realclick=T&tabclick=F&MSK[]=order_id&MSV[]=&orderStatusPayment=all&date_type=order_date&btnDate=0&product_search_type=product_name&find_option=product_no&order_product_name=&order_product_code=&order_product_no=&order_product_text=&order_set_product_no=&layer_order_product_code=&layer_order_product_opt_id=&popup_item_code=&popup_product_code=&payed=&payed_sql_version=&bank_info=&memberType=1&group_no=&isMemAuth=&isBlackList=&isFirstOrder=&isPointfyUsedMember=&shipment_type=all&bunch=&shippedAgain=&shipmentMessage=&delivSeperated=&isReservedOrder=&isSubscriptionOrder=&paystandard=choice&product_total_price1=&product_total_price2=&item_count_start=&item_count_end=&orderPathType=A&search_SaleOpenMarket[]=cafe24&search_SaleOpenMarket[]=mobile&search_SaleOpenMarket[]=mobile_d&search_SaleOpenMarket[]=NCHECKOUT&search_SaleOpenMarket[]=gmarket&search_SaleOpenMarket[]=auction&search_SaleOpenMarket[]=sk11st&search_SaleOpenMarket[]=shopn&search_SaleOpenMarket[]=inpark&search_SaleOpenMarket[]=coupang&search_SaleOpenMarket[]=kakao&search_SaleOpenMarket[]=womanstalk&search_SaleOpenMarket[]=tenten&search_SaleOpenMarket[]=wemake&search_SaleOpenMarket[]=melchi&search_SaleOpenMarket[]=halfclub&search_SaleOpenMarket[]=boribori&search_SaleOpenMarket[]=ogage&search_SaleOpenMarket[]=moongori&search_SaleOpenMarket[]=shopeesg&search_SaleOpenMarket[]=shopeeid&search_SaleOpenMarket[]=shopeemy&search_SaleOpenMarket[]=shopeetw&search_SaleOpenMarket[]=shopeeth&search_SaleOpenMarket[]=shopeeph&search_SaleOpenMarket[]=brich&search_SaleOpenMarket[]=zigzag&search_SaleOpenMarket[]=ably&search_SaleOpenMarket[]=timon&search_SaleOpenMarket[]=musinsa&search_SaleOpenMarket[]=wizwid&search_SaleOpenMarket[]=hottracks&search_SaleOpenMarket[]=akmall&search_SaleOpenMarket[]=daisomall&search_SaleOpenMarket[]=lfmall&search_SaleOpenMarket[]=styleshare&search_SaleOpenMarket[]=aland&search_SaleOpenMarket[]=rakutenkr&search_SaleOpenMarket[]=cjmall&search_SaleOpenMarket[]=lotteon&search_SaleOpenMarket[]=himart&search_SaleOpenMarket[]=tofkof&search_SaleOpenMarket[]=MORUGI&search_SaleOpenMarket[]=11st&mkSaleType=M&mkSaleTypeChg=&inflowPathType=A&inflowPathDetail=0000000000000000000000000000000000&paymethodType=A&paymentMethod[]=cash&paymentMethod[]=card&paymentMethod[]=tcash&paymentMethod[]=icash&paymentMethod[]=cell&paymentMethod[]=deferpay&paymentMethod[]=cvs&paymentMethod[]=point&paymentMethod[]=mileage&paymentMethod[]=deposit&paymentMethod[]=etc&pgListType=A&pgList[]=danal&pgList[]=dacom&pgList[]=payco&pgList[]=paynow&pgList[]=smilepay&pgList[]=eximbay&pgList[]=etc&paymentInfo=&discountMethod=&shop_no_order=1&delvReady=&delvCancel=&orderStatusNotPayCancel=N&orderStatusCancel=N&orderSearchCancelStatus=&orderStatusExchange=N&orderSearchExchangeStatus=&orderStatusReturn=all&orderStatusRefund=N&orderSearchRefundStatus=&orderSearchShipStatus=&orderStatus[]=all&orderStatus[]=N10&orderStatus[]=N20&orderStatus[]=N22&orderStatus[]=N21&orderStatus[]=N30&orderStatus[]=N40&RefundType=&RefundSubType=&sc_id=&second_shipping_company_id=&HopeShipCompanyId=all&post_express_flag=&tabStatus=&paymethod_total_count=&search_invoice_print_flag=all&search_is_escrow_shipping_registered=all&search_print_second_invoice=all&incoming=&is_purchased=&order_fail_code=&isBlackOrder=&start_date=" refdate "&year1=" year "&month1=" monthstr "&day1=" daystr "&start_time=00:00:00&end_date=" refdate "&year2=" year "&month2=" monthstr "&day2=" daystr "&end_time=23:59:59&realclick=T"
		driver.get(url)
		CheckloadComplete(driver)
		sleep,2000
		totalcnt := driver.executeScript("return document.querySelector('p.total strong').innerText")
		;msgbox % totalcnt	
		; save to db...
		url := "https://log.piclick.kr/util/exchange_stat.php?type=return&cnt=" totalcnt "&date=" refdate "&site_id=" site_id
		driver2.get(url) 
		;msgbox % url

		; 전체 환불수
		url := "https://benitomaster.cafe24.com/admin/php/shop1/s_new/order_list_item.php?rows=20&searchSorting=order_desc&isBusanCall=&isChinaCall=&orderCallnum=&cticall=&realclick=T&tabclick=F&MSK[]=order_id&MSV[]=&orderStatusPayment=all&date_type=order_date&btnDate=0&product_search_type=product_name&find_option=product_no&order_product_name=&order_product_code=&order_product_no=&order_product_text=&order_set_product_no=&layer_order_product_code=&layer_order_product_opt_id=&popup_item_code=&popup_product_code=&payed=&payed_sql_version=&bank_info=&memberType=1&group_no=&isMemAuth=&isBlackList=&isFirstOrder=&isPointfyUsedMember=&shipment_type=all&bunch=&shippedAgain=&shipmentMessage=&delivSeperated=&isReservedOrder=&isSubscriptionOrder=&paystandard=choice&product_total_price1=&product_total_price2=&item_count_start=&item_count_end=&orderPathType=A&search_SaleOpenMarket[]=cafe24&search_SaleOpenMarket[]=mobile&search_SaleOpenMarket[]=mobile_d&search_SaleOpenMarket[]=NCHECKOUT&search_SaleOpenMarket[]=gmarket&search_SaleOpenMarket[]=auction&search_SaleOpenMarket[]=sk11st&search_SaleOpenMarket[]=shopn&search_SaleOpenMarket[]=inpark&search_SaleOpenMarket[]=coupang&search_SaleOpenMarket[]=kakao&search_SaleOpenMarket[]=womanstalk&search_SaleOpenMarket[]=tenten&search_SaleOpenMarket[]=wemake&search_SaleOpenMarket[]=melchi&search_SaleOpenMarket[]=halfclub&search_SaleOpenMarket[]=boribori&search_SaleOpenMarket[]=ogage&search_SaleOpenMarket[]=moongori&search_SaleOpenMarket[]=shopeesg&search_SaleOpenMarket[]=shopeeid&search_SaleOpenMarket[]=shopeemy&search_SaleOpenMarket[]=shopeetw&search_SaleOpenMarket[]=shopeeth&search_SaleOpenMarket[]=shopeeph&search_SaleOpenMarket[]=brich&search_SaleOpenMarket[]=zigzag&search_SaleOpenMarket[]=ably&search_SaleOpenMarket[]=timon&search_SaleOpenMarket[]=musinsa&search_SaleOpenMarket[]=wizwid&search_SaleOpenMarket[]=hottracks&search_SaleOpenMarket[]=akmall&search_SaleOpenMarket[]=daisomall&search_SaleOpenMarket[]=lfmall&search_SaleOpenMarket[]=styleshare&search_SaleOpenMarket[]=aland&search_SaleOpenMarket[]=rakutenkr&search_SaleOpenMarket[]=cjmall&search_SaleOpenMarket[]=lotteon&search_SaleOpenMarket[]=himart&search_SaleOpenMarket[]=tofkof&search_SaleOpenMarket[]=MORUGI&search_SaleOpenMarket[]=11st&mkSaleType=M&mkSaleTypeChg=&inflowPathType=A&inflowPathDetail=0000000000000000000000000000000000&paymethodType=A&paymentMethod[]=cash&paymentMethod[]=card&paymentMethod[]=tcash&paymentMethod[]=icash&paymentMethod[]=cell&paymentMethod[]=deferpay&paymentMethod[]=cvs&paymentMethod[]=point&paymentMethod[]=mileage&paymentMethod[]=deposit&paymentMethod[]=etc&pgListType=A&pgList[]=danal&pgList[]=dacom&pgList[]=payco&pgList[]=paynow&pgList[]=smilepay&pgList[]=eximbay&pgList[]=etc&paymentInfo=&discountMethod=&shop_no_order=1&delvReady=&delvCancel=&orderStatusNotPayCancel=N&orderStatusCancel=N&orderSearchCancelStatus=&orderStatusExchange=N&orderSearchExchangeStatus=&orderStatusReturn=N&orderStatusRefund=all&orderSearchRefundStatus=&orderSearchShipStatus=&orderStatus[]=all&orderStatus[]=N10&orderStatus[]=N20&orderStatus[]=N22&orderStatus[]=N21&orderStatus[]=N30&orderStatus[]=N40&RefundType=&RefundSubType=&sc_id=&second_shipping_company_id=&HopeShipCompanyId=all&post_express_flag=&tabStatus=&paymethod_total_count=&search_invoice_print_flag=all&search_is_escrow_shipping_registered=all&search_print_second_invoice=all&incoming=&is_purchased=&order_fail_code=&isBlackOrder=&start_date=" refdate "&year1=" year "&month1=" monthstr "&day1=" daystr "&start_time=00:00:00&end_date=" refdate "&year2=" year "&month2=" monthstr "&day2=" daystr "&end_time=23:59:59&realclick=T"
		driver.get(url)
		CheckloadComplete(driver)
		sleep,2000
		totalcnt := driver.executeScript("return document.querySelector('p.total strong').innerText") 
		;msgbox % totalcnt
		; save to db...
		url := "https://log.piclick.kr/util/exchange_stat.php?type=refund&cnt=" totalcnt "&date=" refdate "&site_id=" site_id
		driver2.get(url)
		;msgbox % url
	}

	return

}

F11::
Reload
;gosub saveSetting
return

GuiClose:
ExitApp
return	
	
