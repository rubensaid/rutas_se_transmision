<?php
require_once('PHPExcel.php');

$id = $_POST['id'];
$id2 = $_POST['id2'];
$points = $_POST['points'];

$kml_file = 'PI.2017-2021.txt';
$file = "Plan.xlsx";
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load($file);
$objWorksheet = $objPHPExcel->setActiveSheetIndexByName('RUTAS');
$setA = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(0, $id)->getValue();
$setB = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(($id2-1), 1)->getValue();

$kml_puntos = json_decode("[".str_replace(Array("(",")"),Array("[","]"),$points)."]");
$kml_p = '';
for($i = 0; $i < count($kml_puntos); $i++) {
	$kml_p .= $kml_puntos[$i][1].",".$kml_puntos[$i][0].",0 ";
}
$kml_content = '
		<Placemark>
			<name>'.$setA.'-'.$setB.'</name>
			<LineString>
				<tessellate>1</tessellate>
				<coordinates>
					'.$kml_p.'
				</coordinates>
			</LineString>
		</Placemark>
';
file_put_contents($kml_file, $kml_content, FILE_APPEND | LOCK_EX);

?>