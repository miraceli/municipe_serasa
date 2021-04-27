<?php

$hostname = "localhost";
$user = "root";
$password = "root";
$database = "db";
$conexao6 = mysqli_connect($hostname, $user, $password, $database);

if(!$conexao6){
    print "Erro ao conectar com o banco de dados.";
}

?>