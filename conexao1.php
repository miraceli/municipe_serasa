<?php

$hostname = "localhost";
$user = "root";
$password = "root";
$database = "db";
$conexao1 = mysqli_connect($hostname, $user, $password, $database);

if(!$conexao1){
    print "Erro ao conectar com o banco de dados.";
}

?>