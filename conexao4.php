<?php

$hostname = "localhost";
$user = "root";
$password = "root";
$database = "db";
$conexao4 = mysqli_connect($hostname, $user, $password, $database);

if(!$conexao4){
    print "Erro ao conectar com o banco de dados.";
}

?>