    
	 <link rel="stylesheet" href="style.css">

<?php
require 'Classes/PHPExcel.php';

$filepath = 'report2.txt';

$date = date('m/d/Y h:i:s a', time());
		
//Variables for chart/s		
	$AccountFieldOnChart=[];
	$HoursFieldOnChart=[];		
		
			
//File Upload
if (empty($_FILES)){
	echo"
	<form method='post' enctype='multipart/form-data' action='index.php'>
		<input type='file' name='excel'>
		<br>
		<button type='submit'>fetch</button>
	</form>
	";
}else{


if (!file_exists($filepath)) {
	$fileName=$_FILES['excel']['name'];
	$file6 = fopen('report.txt', 'w+') or die("File Could Not Be Created");
	fwrite($file6,"\r\n".'Time of analysis is:'.$date."\r\n".'File Opened: '.$fileName.' ...'."\r\n");
	fclose($file6);
}else{
//report files
	$fileName=$_FILES['excel']['name'];
	$file6 = fopen('report.txt', 'a') or die("File Could Not Be Created");
	fwrite($file6,"\r\n".'Time of analysis is:'.$date."\r\n".'File Opened: '.$fileName.' ...'."\r\n"); //Write to txt file
	fclose($file6);
}



//load excel file using PHPExcel's IOFactory
$excel = PHPExcel_IOFactory::load($_FILES['excel']['tmp_name'], 'r');

//set active sheet to first sheet
$excel->setActiveSheetIndex(0);

//load multiple sheets
$CurrentWorkSheetIndex = 0;

//create files needed
$file = fopen('monthly.txt', 'w+') or die("File Could Not Be Created"); 
fclose($file);


foreach ($excel->getWorksheetIterator() as $worksheetNbr =>$worksheet) {

$sheetName = $worksheet->getTitle();
	
	//Test for content
	$ContTest=[];
	for($x='A';$x<'J';$x++)
	{
	$idAccountTestingForCont = $worksheet->getCell($x.'10')->getValue();
	$ContTest += [$x=>$idAccountTestingForCont];
	}
	
	$jointray=join(" ",$ContTest);

		if ($jointray==" ")
		{
			echo 'Worksheet number - ', $worksheetNbr+1, PHP_EOL;
			Print_r($ContTest);
			echo "<br>";
			echo "Sheet Name: ";

			echo $sheetName;
			echo "<hr>";
			echo "Hello, Sheet is Empty!";
			echo "<br>";
			echo "<hr>";
			for($x='A';$x<'J';$x++)
				{
					$idAccountTestingForCont = $worksheet->getCell($x.'10')->getValue();
					echo $idAccountTestingForCont;
					echo "<br>";
					echo "<hr>";
				}
			echo "Hello";
			$ContTest=array();



			}else{
			
			//Employ
			$idEmp = $worksheet->getCell('D8')->getValue();
			//Check if is Employee
			$testidEmp=preg_replace('/\s+/', '', $idEmp);
			$testEmpToLower = strtolower($testidEmp);
			if ($testEmpToLower == 'employee')
				{
					$idEmpPass = $worksheet->getCell('D8')->getValue();		
					//name
					$idN = $worksheet->getCell('D10')->getValue();
					$nme = $worksheet->getCell('E10')->getValue();
					//position
					$idP = $worksheet->getCell('D11')->getValue();
					$poss = $worksheet->getCell('E11')->getValue();
					//department
					$idDe = $worksheet->getCell('D12')->getValue();
					$Dep = $worksheet->getCell('E12')->getValue();
					//salary ref
					$idSal = $worksheet->getCell('H10')->getValue();
					$Sal = $worksheet->getCell('I10')->getValue();
					//unit
					$idUn = $worksheet->getCell('H11')->getValue();
					$Uni = $worksheet->getCell('I11')->getValue();
					//manager
					$idMan = $worksheet->getCell('H12')->getValue();
					$Mana = $worksheet->getCell('I12')->getValue();
					//Pay period
					$idPayLab = $worksheet->getCell('D14')->getValue();
					//from
					$idFr = $worksheet->getCell('D16')->getValue();
					$FrDte = $worksheet->getCell('E16')->getValue();
					$stringDateFrm = \PHPExcel_Style_NumberFormat::toFormattedString($FrDte, 'YYYY-MM-DD');
					//to
					$idTo = $worksheet->getCell('F16')->getValue();	
					$ToDte = $worksheet->getCell('G16')->getValue();
					$stringDateTo = \PHPExcel_Style_NumberFormat::toFormattedString($ToDte, 'YYYY-MM-DD');
					//submited by
					$idSubm = $worksheet->getCell('M8')->getValue();
					$Subm = $worksheet->getCell('O8')->getValue();
					//approved by
					$idApprove = $worksheet->getCell('M14')->getValue();
					$Approve = $worksheet->getCell('O14')->getValue();
				
//=================================================END OF EMPLOYEE DETAILS==============================================

//=========================================LOOP WORK SHEETS==============================================
					echo"<h1 align='center'><p>";
					echo $sheetName.' </p><p>'.'Worksheet number - ', $worksheetNbr+1, PHP_EOL;
					echo "</p></h1>";

//=========================ORDER IN TABLE=======================================
					echo $idEmpPass;
					echo "<hr style='visibility:hidden;'>";
					echo "<table class='responstable' style='float:left;'>";
					echo "
					<!--First Row-->
					<tr>
					<th>$idN</th> <!--Lable for name-->
					<th>$idUn</th> <!--Lable for Unit-->
					<th>$idMan</th> <!--Lable for Manager-->
					</tr><tr>
					<td>$nme</td> <!--name-->
					<td>$Uni</td> <!--Unit-->
					<td>$Mana</td> <!--Manager-->
					</tr>
					";
					echo "</table>";
					//Table 2 Right side top 
					echo "<table class='responstable' align='center'>";
					echo "
					<!--First Row-->
					<tr>
					<th>$idSal</th> <!--Lable for Salary ref-->
					<th>$idP</th> <!--Lable for Position-->
					<th>$idDe</th> <!--Lable for Department-->
					</tr><tr>
					<td>$Sal</td> <!--Salary ref-->
					<td>$poss</td> <!--Position-->
					<td>$Dep</td> <!--Department-->					
					</tr><tr>
					";
					echo "</table>";
					echo "<hr>";
					//Table for dates
					echo $idPayLab;
					echo "<hr style='visibility:hidden;'>";
					echo "<table class='responstable' style='float:left;'>";
					echo "<tr>
					<th>$idFr</th> <!--Lable for from Date-->
					<th>$idTo</th> <!--Lable for to date-->
					</tr>
					<tr>
					<td>$stringDateFrm</td> <!--From Date-->
					<td>$stringDateTo</td> <!--To Date-->
					</tr>
					";
					echo "</table>";
					echo "<hr>";
//=========================END OF ORDER IN TABLE...."TOP HALF"==========================

//======================================Code for Sheet Content==========================
				
					//array for hour in days
					$dataAccount = [];
					$dataAccountCode = [];
					$dataMon = [];
					$dataTue = [];
					$dataWed = [];
					$dataThur = [];
					$dataFri = [];
					$dataSat = [];
					$dataSun = [];
					$dataMon2 = [];
					$dataTue2 = [];
					$dataWed2 = [];
					$dataThur2 = [];
					$dataFri2 = [];
					$dataSat2 = [];
					$dataSun2 = [];
					echo "Accounts";
					echo "<table class='responstable' border='1' style='border:solid;border-collapse: collapse;'>";
					//first row of data
					$i = 21;

					//loop until the end of data cell "Notes and Remarks"
					$esti=$worksheet->getCell('I'.$i)->getValue();
					$testMatch=preg_replace('/\s+/', '', $esti);
					do{
						//content check
						$idAccountCheck = $worksheet->getCell('D'.$i)->getValue();
						$testMatch4=preg_replace('/\s+/', '', $idAccountCheck);
						
						if($testMatch4==""){
						$idAccount=="No Account";
						}else{
						
						
						//get cells value
						$idAccount = $worksheet->getCell('D'.$i)->getValue();
						$dataAccount += [$i=>$idAccount];
						$idAccCde = $worksheet->getCell('I'.$i)->getValue();
						$dataAccountCode += [$i=>$idAccCde];
						$idMon = $worksheet->getCell('J'.$i)->getCalculatedValue();
						$dataMon += [$i=>$idMon];
						$idTue = $worksheet->getCell('K'.$i)->getCalculatedValue();
						$dataTue += [$i=>$idTue];
						$idWed = $worksheet->getCell('L'.$i)->getCalculatedValue();
						$dataWed += [$i=>$idWed];
						$idThur = $worksheet->getCell('M'.$i)->getCalculatedValue();
						$dataThur += [$i=>$idThur];
						$idFri = $worksheet->getCell('N'.$i)->getCalculatedValue();
						$dataFri += [$i=>$idFri];
						$idSat = $worksheet->getCell('O'.$i)->getCalculatedValue();
						$dataSat += [$i=>$idSat];	
						$idSun = $worksheet->getCell('P'.$i)->getCalculatedValue();
						$dataSun += [$i=>$idSun];	

						$headOfTableAcc=preg_replace('/\s+/', '', $idAccount);
						if($headOfTableAcc=='AccountDescription'){
							echo "
							<tr>
							<th>$idAccount</th>
							<th>$idAccCde</th>
							<th>$idMon</th>
							<th>$idTue</th>
							<th>$idWed</th>
							<th>$idThur</th>
							<th>$idFri</th>
							<th>$idSat</th>
							<th>$idSun</th>		
							</tr>
							";
						}else{
							echo "
							<tr>
							<td>$idAccount</td>
							<td>$idAccCde</td>
							<td>$idMon</td>
							<td>$idTue</td>
							<td>$idWed</td>
							<td>$idThur</td>
							<td>$idFri</td>
							<td>$idSat</td>
							<td>$idSun</td>		
							</tr>
							";
							}
						}
						//row pointer
						$i++;
						
						
						$esti=$worksheet->getCell('I'.$i)->getValue(); //stop looping
						$esti2=$worksheet->getCell('G'.$i)->getValue();
			
						${'Mothly'.$i.$sheetName}=[];
						for($x='A';$x<'Q';$x++){
							$idAccountTestingForCont = $worksheet->getCell($x.$i)->getCalculatedValue();
							${'Mothly'.$i.$sheetName}+=[$x=>$idAccountTestingForCont];
						}
						
			
						if ($i>50){
							$esti3=$worksheet->getCell('A50')->getValue();		
						}else{
							$esti3 = "nothing";
						}
				
						$testMatch=preg_replace('/\s+/', '', $esti);
						$testMatch2=preg_replace('/\s+/', '', $esti2);
						$testMatch3=preg_replace('/\s+/', '', $esti3);
		
						if($testMatch=="TotalHours" || $testMatch2=="TotalHours" || $testMatch3==""){
							$endOfL = "Yes";
						}else {
							$endOfL = "No";
						}
		
					}while( $endOfL == "No" );	
			
					echo "</table>";

//=================================Code for approved/Submit by=================================

					echo "<hr>";
					//Table for approve
					echo "<table class='responstable'>";
					echo "
					<tr>
					<th>$idApprove</th> <!--Lable for from Date-->
					<th>$idSubm</th> <!--Lable for to date-->
					</tr><tr>
					<td>$Approve</td> <!--From Date-->
					<td>$Subm</td> <!--To Date-->
					</tr>
					";
					echo "</table>";

//====================================Code for Analysis========================================

					echo "<hr>";
					echo "<h1 align='center'>Analysis </h1>";
					echo "<h4>Whole Day. </h4>";
			
					//date with hours
					$plus_a_day = date ("Y-M-d D", strtotime($stringDateFrm ."+0 days"));
					$plus_a_day1 = date ("Y-M-d D", strtotime($stringDateFrm ."+1 days"));
					$plus_a_day2 = date ("Y-M-d D", strtotime($stringDateFrm ."+2 days"));
					$plus_a_day3 = date ("Y-M-d D", strtotime($stringDateFrm ."+3 days"));
					$plus_a_day4 = date ("Y-M-d D", strtotime($stringDateFrm ."+4 days"));
					$plus_a_day5 = date ("Y-M-d D", strtotime($stringDateFrm ."+5 days"));
					$plus_a_day6 = date ("Y-M-d D", strtotime($stringDateFrm ."+6 days"));

					//date with hours
					echo "<table class='responstable'>";
					echo "<tr><th>Day</th><th>Hours</th></tr>";
				
					echo "<tr>
					<td>$plus_a_day </td><td> Total Hours Spent: ".array_sum($dataMon);
					echo "</td></tr><tr>
					<td>$plus_a_day1 </td><td> Total Hours Spent: ".array_sum($dataTue);
					echo "</td></tr><tr>
					<td>$plus_a_day2 </td><td> Total Hours Spent: ".array_sum($dataWed);
					echo "</td></tr><tr>
					<td>$plus_a_day3 </td><td> Total Hours Spent: ".array_sum($dataThur);
					echo "</td></tr><tr>
					<td>$plus_a_day4 </td><td> Total Hours Spent:".array_sum($dataFri);
					echo "</td></tr><tr>
					<td>$plus_a_day5 </td><td> Total Hours Spent: ".array_sum($dataSat);
					echo "</td></tr><tr>
					<td>$plus_a_day6 </td><td> Total Hours Spent: ".array_sum($dataSun);
					echo "</td></tr>
					";
					echo "</table>";

					echo "<h4>Whole week.</h4>";
					echo "<table class='responstable' border='' style='border:solid;border-collapse: collapse;'>";
		
					$AccountFieldOnChartBar=[];
				// loop for week calculation
				foreach($dataAccount as $i => $item){ 
				
					$headOfTable=preg_replace('/\s+/', '', $dataAccount[$i]);

						if($headOfTable=='AccountDescription'){		
						echo "<tr><th>";
						echo $dataAccount[$i]; //accont name
						echo "</th><th>";					
						echo $dataAccountCode[$i]; //account code
						echo "</th><th>";
						echo"Time";
						echo "</th><th>";
						echo"Date From";
						echo "</th><th>";
						echo"Date To";
						echo "</th></tr>";
					}else{
						echo "<tr> <td>";	
						echo $dataAccount[$i]; //accont name
						echo "</td><td>";
						echo $dataAccountCode[$i]; //account code
						echo "</td><td>";
						$dataMon2 += [$i=>$dataMon[$i]];
						$dataTue2 += [$i=>$dataTue[$i]];
						$dataWed2 += [$i=>$dataWed[$i]];					
						$dataThur2 += [$i=>$dataThur[$i]];
						$dataFri2 += [$i=>$dataFri[$i]];
						$dataSat2 += [$i=>$dataSat[$i]];
						$dataSun2 += [$i=>$dataSun[$i]];
						$totalWeekHori=$dataMon2[$i]+$dataTue2[$i]+$dataWed2[$i]+$dataThur2[$i]+$dataFri2[$i]+$dataSat2[$i]+$dataSun2[$i];
						//graph
						$AccountFieldOnChartBar+=[$dataAccount[$i]=>$totalWeekHori];
						//graph end
						if ($totalWeekHori > 0) {
							echo $totalWeekHori." Hours";
							echo "</td><td>";
							echo $stringDateFrm; 
							echo "</td><td>";
							echo $stringDateTo;
						}else{
							echo $totalWeekHori;
							echo "</td><td>";
							echo $stringDateFrm; 
							echo "</td><td>";
							echo $stringDateTo;
						}
					}
				echo "</td></tr>";
				}
				echo "</table>";
				
				//Chart
				include 'bar.php';			

				echo "<hr>";
				$ContTest=array();//Flush array for content test	
			}else{
				
									$dataAccount = [];
					$dataAccountCode = [];
					$dataMon = [];
					$dataTue = [];
					$dataWed = [];
					$dataThur = [];
					$dataFri = [];
					$dataSat = [];
					$dataSun = [];
					$dataMon2 = [];
					$dataTue2 = [];
					$dataWed2 = [];
					$dataThur2 = [];
					$dataFri2 = [];
					$dataSat2 = [];
					$dataSun2 = [];
				
									$i = 21;

					//loop until the end of data cell "Notes and Remarks"
					$esti=$worksheet->getCell('I'.$i)->getValue();
					$testMatch=preg_replace('/\s+/', '', $esti);
					do{
						//get cells value
						$idAccount = $worksheet->getCell('D'.$i)->getValue();
						$dataAccount += [$i=>$idAccount];
						$idAccCde = $worksheet->getCell('I'.$i)->getValue();
						$dataAccountCode += [$i=>$idAccCde];
						$idMon = $worksheet->getCell('J'.$i)->getCalculatedValue();
						$dataMon += [$i=>$idMon];
						$idTue = $worksheet->getCell('K'.$i)->getCalculatedValue();
						$dataTue += [$i=>$idTue];
						$idWed = $worksheet->getCell('L'.$i)->getCalculatedValue();
						$dataWed += [$i=>$idWed];
						$idThur = $worksheet->getCell('M'.$i)->getCalculatedValue();
						$dataThur += [$i=>$idThur];
						$idFri = $worksheet->getCell('N'.$i)->getCalculatedValue();
						$dataFri += [$i=>$idFri];
						$idSat = $worksheet->getCell('O'.$i)->getCalculatedValue();
						$dataSat += [$i=>$idSat];	
						$idSun = $worksheet->getCell('P'.$i)->getCalculatedValue();
						$dataSun += [$i=>$idSun];	
						echo "
						<tr>
						<td>$idAccount</td>
						<td>$idAccCde</td>
						<td>$idMon</td>
						<td>$idTue</td>
						<td>$idWed</td>
						<td>$idThur</td>
						<td>$idFri</td>
						<td>$idSat</td>
						<td>$idSun</td>		
						</tr>
						";
						//row pointer
						$i++;
						
						
						$esti=$worksheet->getCell('I'.$i)->getValue(); //stop looping
						$esti2=$worksheet->getCell('G'.$i)->getValue();
			
						${'Mothly'.$i.$sheetName}=[];
						for($x='A';$x<'Q';$x++)
							{
								$idAccountTestingForCont = $worksheet->getCell($x.$i)->getCalculatedValue();
								${'Mothly'.$i.$sheetName}+=[$x=>$idAccountTestingForCont];
							}
						
			
						if ($i>50){
							$esti3=$worksheet->getCell('A50')->getValue();		
						}else{
							$esti3 = "nothing";
						}
				
						$testMatch=preg_replace('/\s+/', '', $esti);
						$testMatch2=preg_replace('/\s+/', '', $esti2);
						$testMatch3=preg_replace('/\s+/', '', $esti3);
		
						if($testMatch=="TotalHours" || $testMatch2=="TotalHours" || $testMatch3==""){
							$endOfL = "Yes";
						}else {
							$endOfL = "No";
						}
		
					}while( $endOfL == "No" );	
				
				echo "<br>";
				$idEmpPass='Error Loading Worksheet. Please Check WorkBook!';
				echo $idEmpPass;
				echo "<br>";
				echo "Error found in...";
				echo "<br>";
				echo "Sheet Name: ";
				$sheetName = $worksheet->getTitle();
				echo $sheetName;
				echo"<br>";
				echo "Sheet number: ";
				echo $worksheetNbr+1;
				echo "<br>";
				echo "<hr>";
			}
		}	
		unset($worksheet);
		//Monthly Calculation
	}
	
//Outside loop==============================TAKE ME TOO====================
	
	$TotalsheetName=$excel->getSheetNames();

	foreach($TotalsheetName as $TotalsheetName2 => $dilo){
		$i = 21;
		do{
			$i++;
			$trimmedArray = array_map('trim', ${'Mothly'.$i.$dilo});//remove space
			$rr=array_filter($trimmedArray, 'strlen');//reomve empty elements in array						

			//write to file
			$withComma = implode("\t", $rr);				
			$const=0;
			$file2 = fopen('monthly.txt','a'); //a to append w to write
			fwrite($file2,$withComma . "\r\n");
			fclose($file2);					

			//end loop here
			$esti=$excel->getActiveSheet()->getCell('I'.$i)->getValue();
			$testMatch=preg_replace('/\s+/', '', $esti);

			if($testMatch=="TotalHours"){
				$endOfL = "Yes";
			}else{
				$endOfL = "No";
			}
		}while( $endOfL == "No" );	
	}



	//Remove empty spaces from file
	file_put_contents('newFile.txt',
		preg_replace(
			'~[\r\n]+~',
			"\r\n",
        trim(file_get_contents('monthly.txt'))
		)
	);



	// read file into array
	$arrayF = file('newFile.txt');
	// new array to store results
	$new_array = array();
	// loop through array
	foreach ($arrayF as $line) {
		// explode the line on tab. Note double quotes around \t are mandatory
		$line_array = explode("\t", $line);
		// set first element to the new array
		$calculate_array[] = $line_array[0];
	}
	$calculate_arrayNoDuplicate = array_unique($calculate_array);

	echo "<h1 align='center'>Analysis of Accounts for The whole Book</h1>";
	echo "<hr style='visibility:hidden;'>";
	//Place Table Here (whole book)
	echo "<table class='responstable'>";
	echo "<tr><th>Account Description</th><th>Hours Worked</th></tr>";

	foreach($calculate_arrayNoDuplicate as $finish => $m){
		$RemoSpace=preg_replace('/\s+/', '', $m);
		if($RemoSpace=='TotalHours'){
			continue;
		}else{
			$searchthis = $m; //Match search
			$matches = array();
			$handle = @fopen("newFile.txt", "r");
			if ($handle){
				while (!feof($handle)){ 
				//===========================Check for matches in file
					$buffer = fgets($handle);
					if(strpos($buffer, $searchthis) !== FALSE)
						$matches[] = $buffer; 
				}
				fclose($handle); //close file
			}

			//=========================Monyhly Calculation====================
			$trimmedArray = array_map('trim', $matches);//remove space
			$ry=array_filter($trimmedArray, 'strlen');//reomve empty elements in array 				
			$file2=fopen('calculate.txt','w+'); //file for calculating
			foreach($ry as $call => $t){
				$file2 = fopen('calculate.txt','a'); //a to append w to write
				fwrite($file2,$t.'	0	'. "\r\n"); //Write to txt file
			}					
			fclose($file2);
			$textCnt  = "calculate.txt"; //file to calculate
			$contents = file_get_contents($textCnt); 
			$arrfields = explode("\t", $contents);//read file separated by comma/by tab
			$ad=(float)0;
			//check if it is Account code or time
			foreach($arrfields as $fields=>$t2) {
				$RemoSpace2=preg_replace('/\s+/', '', $t2);
				if ($t2>25){
					continue;
				}else{
					$t2F=(float)$t2; //Change the return to float
					$ad+=$t2F;
				}
			}
	
			echo "<tr><td>".$arrfields[0]."</td>"; //Display the account being calculated
			echo "<td>".$ad."</td></tr>";	//Calculation for whole month					
			$file5 = fopen('report.txt','a'); //a to append w to write
				fwrite($file5,$nme.'. Account: '.$arrfields[0] .' - '.$ad.' Hours'. "\r\n"); //Write to txt file
			fclose($file5);
			//calculations for chart
			$AccountFieldOnChart+=[$arrfields[0]=>$ad];
		}
	}
	echo "</table>";

	echo "<br><strong><u>*If a sheet failed to load analysis is false*</u></strong>";
	//back up file
	file_put_contents('report2.txt',
		preg_replace(
			'~[\r\n]+~',
			"\r\n",
        trim(file_get_contents('report.txt'),FILE_APPEND | LOCK_EX)
		)
	);
	//chart for whole book
	include 'pie.php';
}		
?>