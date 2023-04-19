<?php
// Conexão com o banco de dados Oracle
$tns = "(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = host)(PORT = port))) (CONNECT_DATA = (SERVICE_NAME = service_name)))"; 
// substitua pelos valores corretos no host, porta, service name acima e login e senha abaixo
$db_username = "usuario"; // substitua pelo valor correto
$db_password = "senha"; // substitua pelo valor correto

try {
    $conn = new PDO("oci:dbname=".$tns, $db_username, $db_password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Leitura do arquivo Excel
    $arquivo_excel = 'c:/pastapublica/bra.xls';
    $planilha = PHPExcel_IOFactory::load($arquivo_excel);
    $conteudo = $planilha->getActiveSheet()->getCell('B2')->getValue();
    $material = $planilha->getActiveSheet()->getCell('D9')->getValue();

    // Inserção no banco de dados
    $stmt = $conn->prepare("INSERT INTO usuario (conteudo, material) VALUES (:conteudo, :material)");
    $stmt->bindParam(':conteudo', $conteudo);
    $stmt->bindParam(':material', $material);
    $stmt->execute();

    echo "Inserção realizada com sucesso!";
}
catch(PDOException $e) {
    echo "Erro ao inserir dados no banco de dados: " . $e->getMessage();
}

// Encerramento da conexão com o banco de dados
$conn = null;
?>
