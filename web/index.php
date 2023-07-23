<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel usando PHP</title>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap-theme.min.css">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js"></script>
</head>
<body>
            
<?php
require_once 'PHPExcel/Classes/PHPExcel.php';
$archivo = "ListaPersonal.xlsx";
$inputFileType = PHPExcel_IOFactory::identify($archivo);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objPHPExcel = $objReader->load($archivo);
$sheet = $objPHPExcel->getSheet(0); 
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();


$servidor = "localhost:3307";
$username="root";
$password="";
$database="listapersonal";

$mysqli = new mysqli($servidor, $username, $password, $database);

if ($mysqli->connect_errno) {
    echo "Falló la conexión con MySQL: (" . $mysqli->connect_errno . ") " . $mysqli->connect_error;
	die();
}


	if( !$mysqli->query("DELETE FROM hoja1") ) {
		echo "Falló EL BORRADO DE LOS DATOS: (" . $mysqli->errno . ") " . $mysqli->error;
		die();
	}

	echo "<hr>";	
	echo "Carga los Datos a la Base de Datos ";	
	echo "<hr>";	

$num=0;
for ($row = 2; $row <= $highestRow; $row++){ $num++;

	echo $num;
	echo " ";
	echo $sheet->getCell("A".$row)->getValue();
	echo " ";
	echo $sheet->getCell("B".$row)->getValue();
	echo " ";
	echo $sheet->getCell("C".$row)->getValue();
	echo " ";
	echo $sheet->getCell("D".$row)->getValue();
	echo "<br>";

	$a = $sheet->getCell("A".$row)->getValue();
	$b = $sheet->getCell("B".$row)->getValue();
	$c = $sheet->getCell("C".$row)->getValue();
	$d = $sheet->getCell("D".$row)->getValue();
	
	if( !$mysqli->query("INSERT INTO hoja1 (id, valora, valorb, valorc, valord) VALUES ('$num', '$a', '$b', '$c', '$d')") ) {
		echo "Falló Insercion de la tabla: (" . $mysqli->errno . ") " . $mysqli->error;
	}
	
}

	echo "<br>";
	echo "OK";
	
	echo "<hr>";	
	echo "Muestra la Consulta de la Base de Datos ";	
	echo "<hr>";	
	
$resultado = $mysqli->query("select * FROM hoja1");

//$resultado = $mysqli->query("select id, valora, valorb, valorc, valord FROM hoja1");

echo "ejecutar Select";
echo "<br>";

echo "select id, valora, valorb, valorc, valord FROM hoja1";
echo "<br>";

echo "Orden inverso...\n";
echo "<br>";

for ($num_fila = $resultado->num_rows - 1; $num_fila >= 0; $num_fila--) {
    $resultado->data_seek($num_fila);
    $fila = $resultado->fetch_assoc();
    echo " id = " . $fila['id'] . " " . $fila['valora'] . " " . $fila['valorb'] . " " . $fila['valorc'] . " " . $fila['valord'] . " " . "\n" ;
	echo "<br>";
}

echo "<br>";
echo "Orden del conjunto de resultados...\n";
echo "<br>";

$resultado->data_seek(0);
while ($fila = $resultado->fetch_assoc()) {
    echo " id = " . $fila['id'] . " " . $fila['valora'] . " " . $fila['valorb'] . " " . $fila['valorc'] . " " . $fila['valord'] . " " . "\n" ;
	echo "<br>";
}


echo "<br>";
echo "OK";




/*
CREATE TABLE hoja1 (
	`id` INT(11) NOT NULL,
	`valora` VARCHAR(50) NULL DEFAULT NULL COLLATE 'utf8_general_ci',
	`valorb` VARCHAR(50) NULL DEFAULT NULL COLLATE 'utf8_general_ci',
	`valorc` VARCHAR(50) NULL DEFAULT NULL COLLATE 'utf8_general_ci',
	`valord` VARCHAR(50) NULL DEFAULT NULL COLLATE 'utf8_general_ci',
	PRIMARY KEY (`id`) USING BTREE
)
COLLATE='utf8_general_ci'
ENGINE=InnoDB
;
*/

$mysqli->close();	

?>
<br>
<br>
<a href="ver.php">Ver Hoja De Excel</a>
</body>
</html>
