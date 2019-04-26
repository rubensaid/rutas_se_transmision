<?php
require_once('PHPExcel.php');

if(!isset($_GET['step'])) $step = 0; else $step = $_GET['step'];
if(!isset($_GET['id'])) $id = 2; else $id = $_GET['id'];
if($step == 2) $id2_b = 2; else $id2_b = 2; /* First value could be starting row */
if($id+1 > $id2_b) $id2_b = $id+1;
if(!isset($_GET['id2'])) $id2 = $id2_b; else $id2 = $_GET['id2'];

print '
<html>
<head>
	<title>Generador de Rutas | Paso '.$step.' | ID1: '.$id.' | ID2: '.$id2.'</title>
	<script src="jquery-2.1.3.js"></script>
	<script type="text/javascript" src="https://maps.google.com/maps/api/js?v=quarterly&amp;libraries=geometry,visualization,places&amp;key=TU_API_GOOGLE_AQUI"></script>
	<script src="jscoord-1.1.1.js"></script>
	<script>
		function convertUTM(cord1, cord2) {
			var utm = new UTMRef(cord1, cord2, "E", 18);
			var latlon = utm.toLatLng();
			$("input[name=latlng]").val(latlon.lat+"y"+latlon.lng);
			$("form[name=pasar]").submit();
		}

		var directionsService = new google.maps.DirectionsService();		

		function calcDist(lat1, lng1, lat2, lng2) {	
			largoLT = 0; largopuntos = "";
			cortoLT = 0; cortopuntos = "";
			finalLT = 0; finalpuntos = "";
			setA = new google.maps.LatLng(lat1, lng1);
			setB = new google.maps.LatLng(lat2, lng2);		
			request = {
				origin: setA,
				destination: setB,
				travelMode: google.maps.TravelMode.DRIVING
			};
			directionsService.route(request, function(response, status) {
				if (status == google.maps.DirectionsStatus.OK) {
					largoLT = Math.round(response.routes[0].legs[0].distance.value/1000*10)/10;
					largopuntos = response.routes[0].overview_path;
				}
			});
			request = {
				origin: setB,
				destination: setA,
				travelMode: google.maps.TravelMode.DRIVING
			};
			directionsService.route(request, function(response, status) {
				if (status == google.maps.DirectionsStatus.OK) {
					cortoLT = Math.round(response.routes[0].legs[0].distance.value/1000*10)/10;
					cortopuntos = response.routes[0].overview_path;
				}
			});
setTimeout(function(){
    console.log("Durmiendo...");
			if(cortoLT < largoLT) {
				finalLT = cortoLT;
				finalpuntos = cortopuntos;
			} else {
				finalLT = largoLT;
				finalpuntos = largopuntos;
			}
			if(finalLT > 15) { //en km
				$("input[name=dist]").val("NR");
				$("form[name=pasar]").submit();
			} else {
				$("input[name=dist]").val(finalLT);
				var convertedArray = [];
				for(var i = 0; i < finalpuntos.length; ++i)
				{
					convertedArray.push(finalpuntos[i]);
				}				
				$.post("puntos.php",{id:'.$id.', id2:'.$id2.', points:convertedArray.join()}, function() {
					$("form[name=pasar]").submit();
				});
			}
}, 2800);
		}
	</script>
</head>
<body>
';

$file = "Plan.xlsx";
$kml_file = 'Rutas.txt';
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load($file);

/*****************************************/
/*                 PASO 0                */
/*       LECTURA DE NOMBRES DE SETs      */
/*          Y CREACIÓN DE MATRIZ         */
/*****************************************/

if($step == 0) {

if(file_exists($kml_file)) unlink($kml_file);
if(file_exists(str_replace(".txt",".kml",$kml_file))) unlink(str_replace(".txt",".kml",$kml_file));
fopen($kml_file, "w");
$kml_content = '<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://earth.google.com/kml/2.1">
<Document>
	<name>PI 2017-2021</name>
	<Folder>
		<name>SET</name>
		<open>1</open>
';
file_put_contents($kml_file, $kml_content, FILE_APPEND | LOCK_EX);

$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('NODOS');
$num_nod = $objWorksheet->getHighestRow();
$SETs = Array();
for($i=2; $i<=$num_nod; $i++) {
	$SETs[] = $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getValue();
}
for($i=0; $i<count($SETs); $i++) {
	$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('RUTAS');
	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, ($i+2), $SETs[$i]);
	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(($i+1), 1, $SETs[$i]);
}
$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
$objWriter->save("Plan.xlsx");

header("Location: ?step=1");

}

/*****************************************/
/*                 PASO 1                */
/*        CONVERSIÓN A COORDENADAS       */
/*                DECIMALES              */
/*****************************************/

if($step == 1) {

$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('NODOS');
$num_nod = $objWorksheet->getHighestRow();

if($id > $num_nod) {
$kml_content = '
	</Folder>
	<Folder>
		<name>LT</name>
		<open>1</open>
';
file_put_contents($kml_file, $kml_content, FILE_APPEND | LOCK_EX);
	header("Location: ?step=2");
}

$set = $objPHPExcel->getActiveSheet()->getCell('A'.$id)->getValue();
$utm_x = $objPHPExcel->getActiveSheet()->getCell('B'.$id)->getValue();
$utm_y = $objPHPExcel->getActiveSheet()->getCell('C'.$id)->getValue();
if($objPHPExcel->getActiveSheet()->getCell('D'.$id)->getValue() > 0) {
	header("Location: ?step=1&id=".($id+1));
} else {
	if(isset($_GET['latlng'])) {
		$latlng = explode("y",$_GET['latlng']);
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$id, $latlng[0]);
		$objPHPExcel->getActiveSheet()->setCellValue('E'.$id, $latlng[1]);
		$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
		$objWriter->save("Plan.xlsx");
$kml_content = '
		<Placemark>
			<name>'.$set.'</name>
			<Point>
				<gx:drawOrder>1</gx:drawOrder>
				<coordinates>'.$latlng[1].','.$latlng[0].',0</coordinates>
			</Point>
		</Placemark>
';
file_put_contents($kml_file, $kml_content, FILE_APPEND | LOCK_EX);
		header("Location: ?step=1&id=".($id+1));
	} else {
		print '
			<h1>Paso 1: Convertiendo a coordenadas decimales SET '.$set.'</h1>
			<form action="'.$_SERVER['PHP_SELF'].'" method="get" name="pasar">
				Paso: <input name="step" value="'.$step.'">
				ID: <input name="id" value="'.$id.'">
				Lat/Long: <input name="latlng">
			</form>
			<script>convertUTM('.$utm_x.', '.$utm_y.')</script>
		';
	}
}

}

/*****************************************/
/*                 PASO 2                */
/*      CALCULO DE DISTANCIAS ENTRE      */
/*                  SET's                */
/*****************************************/

if($step == 2) {
$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('NODOS');
$num_nod = $objWorksheet->getHighestRow();
$SETs = Array();
for($i=2; $i<=$num_nod; $i++) {
	$SETs[$i] = Array(
		"lat" => $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue(),
		"lng" => $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue()
	);
}
$num_nod--;
$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('RUTAS');
$setA = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(0, $id)->getValue();
$setB = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(($id2-1), 1)->getValue();
if($objPHPExcel->getActiveSheet()->getCellByColumnAndRow(($id2-1), $id)->getValue() != "" or ($id == $id2)) { //or (!is_numeric(substr($setA,1,strlen($setA))) && !is_numeric(substr($setB,1,strlen($setB))))
	echo $setA." con ".$setB;
	if($id2 > $num_nod) {
		if($id > $num_nod) {
			header("Location: ?step=3");
		} else {
			header("Location: ?step=2&id=".($id+1)."&id2=".$id2_b); //($id+1)
		}
	} else {
		header("Location: ?step=2&id=".$id."&id2=".($id2+1));
	}
} else {
	if($id == $id2) {	
/*		$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(($id2-1), $id, 0);
		$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
		$objWriter->save("Plan.xlsx");*/
		if($id2 > $num_nod) {
			if($id > $num_nod) {
				header("Location: ?step=3");
			} else {
				header("Location: ?step=2&id=".($id+1)."&id2=".$id2_b); //($id+1)
			}
		} else {
			header("Location: ?step=2&id=".$id."&id2=".($id2+1));
		}
	} else if(isset($_GET['dist'])) {
		$dist = $_GET['dist'];
		if($dist != "NR") {
			$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(($id2-1), $id, $dist);
			$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
			$objWriter->save("Plan.xlsx");
		}
		if($id2 > $num_nod) {
			if($id > $num_nod) {
				header("Location: ?step=3");
			} else {
				header("Location: ?step=2&id=".($id+1)."&id2=".$id2_b); //($id+1)
			}
		} else {
			header("Location: ?step=2&id=".$id."&id2=".($id2+1));
		}
	} else {
		print '
			<h1>Paso 2: Calculando rutas entre SETs '.$setA.' y '.$setB.'</h1>
			<form action="'.$_SERVER['PHP_SELF'].'" method="get" name="pasar">
				Paso: <input name="step" value="'.$step.'">
				ID: <input name="id" value="'.$id.'">
				ID2: <input name="id2" value="'.$id2.'">
				Distancia: <input name="dist">
			</form>
			<script>calcDist('.$SETs[$id]['lat'].', '.$SETs[$id]['lng'].', '.$SETs[$id2]['lat'].', '.$SETs[$id2]['lng'].')</script>
		';
	}
}

}

/*****************************************/
/*                 PASO 3                */
/*          RENOMBRANDO ARCHIVOS Y       */
/*         PREPARANDO PRESENTACIÓN       */
/*****************************************/

if($step == 3) {
	$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('RUTAS');
	$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
	$objWriter->save("Plan.xlsx");
$kml_content = '
	</Folder>
</Document>
</kml>
';
file_put_contents($kml_file, $kml_content, FILE_APPEND | LOCK_EX);
rename($kml_file, str_replace(".txt",".kml",$kml_file));
	//header("Refresh: 1; URL=Plan.xlsx");
	//header("Refresh: 1; URL=".str_replace(".txt",".kml",$kml_file));
	print '
		<h1>Paso 3: Trabajo completado</h1>
		<a href="Plan.xlsx">Descargar archivo XLSX con distancias entre SETs</a>
		<br />
		<a href="'.str_replace(".txt",".kml",$kml_file).'">Descargar archivo KML con las SETs y enlaces dibujados en Google Earth</a>
	';
}

print '
</body>
</html>
';

?>