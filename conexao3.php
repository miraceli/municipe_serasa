<?php

$hostname = "localhost";
$user = "root";
$password = "root";
$database = "db";
$conexao3 = mysqli_connect($hostname, $user, $password, $database);

if(!$conexao3){
    print "Erro ao conectar com o banco de dados.";
}

?>