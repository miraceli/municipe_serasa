<?php

$hostname = "localhost";
$user = "root";
$password = "root";
$database = "db";
$conexao5 = mysqli_connect($hostname, $user, $password, $database);

if(!$conexao5){
    print "Erro ao conectar com o banco de dados.";
}

?>