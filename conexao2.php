<?php

$hostname = "localhost";
$user = "root";
$password = "root";
$database = "db";
$conexao2 = mysqli_connect($hostname, $user, $password, $database);

if(!$conexao2){
    print "Erro ao conectar com o banco de dados.";
}

?>