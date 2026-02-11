<?php
// Archivo TRM de respaldo - valores por defecto
header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

$trm_data = [
    "usd" => 3921.58,
    "eur" => 4000.00,
    "fecha" => date("Y-m-d H:i:s")
];

echo json_encode($trm_data);
?>
