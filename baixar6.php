<?php
	session_start();
	include_once('conexao6.php');
?>
<!DOCTYPE html>
<html lang="pt-br">
	<head>
		<meta charset="utf-8">
		<title>Munícipes Serasa</title>
	<head>
	<body>
		<?php

		$hoje = date('d/m/Y');
		$mes = date('m');
		$ano = date('Y');
		$nome = 'SFS - Serasa - Anotações - ';
		$extensao = '.xls';
		$arquivo = $nome . $mes . $ano .$extensao;
        //$arquivo = 'BC - Serasa - Anotações - 042021.xls';
        
        $html = '';
		$html .= '<p><strong>CONCESSÃO SÃO FRANCISCO DO SUL - MUNÍCIPES COM ANOTAÇÕES ATIVAS NA SERASA</strong></p>';
		$html .= '<p><strong>Atualizado em: '.$hoje.'</strong></p></br></br>';
		$html .= '<table border="1">';		
		$html .= '<tr>';
		$html .= "<td style='text-align:center;vertical-align:middle'><b>Nome Munícipe</b></td>";
		$html .= "<td style='text-align:center;vertical-align:middle'><b>Qtd. Imóveis</b></td>";
		$html .= "<td style='text-align:center;vertical-align:middle'><b>Tarifa</b></td>";
		$html .= "<td style='text-align:center;vertical-align:middle'><b>Acréscimo</b></td>";
        $html .= "<td style='text-align:center;vertical-align:middle'><b>Total</b></td>";
		$html .= '</tr>';
		
        $sql = "
		SELECT Municipe.NomeMunicipe, COUNT(DISTINCT LoteCobrancaSerasa.CodImovel) AS QtdImoveis, CONCAT('R$ ',(REPLACE((SUM(Cobranca.VlrTarifa)), '.', ','))) AS Tarifa, CONCAT('R$ ',(REPLACE((SUM(Cobranca.VlrAcrescimo)), '.', ','))) AS Acrescimo, CONCAT('R$ ',(REPLACE((SUM(Cobranca.VlrTarifa + Cobranca.VlrAcrescimo)), '.', ','))) AS Total
		FROM LoteCobrancaSerasa 
		INNER JOIN CobrancaSerasa ON (LoteCobrancaSerasa.NrDocto = CobrancaSerasa.NrConv) 
		INNER JOIN Cobranca ON (CobrancaSerasa.NrCobrancaOrigem = Cobranca.NrCobranca) 
		INNER JOIN Municipe ON (LoteCobrancaSerasa.CodMunicipe = Municipe.CodMunicipe)
		WHERE CobrancaSerasa.Situacao = 'A'
		GROUP BY LoteCobrancaSerasa.CodMunicipe
		ORDER BY Municipe.NomeMunicipe;
";
		$consulta6 = mysqli_query($conexao6 , $sql);
		
		while($exibirRegistros = mysqli_fetch_array($consulta6)){
			$html .= '<tr>';
			$html .= "<td style='text-align:left;vertical-align:middle'>".$exibirRegistros[0].'</td>';
			$html .= "<td style='text-align:center;vertical-align:middle'>".$exibirRegistros[1].'</td>';
			$html .= "<td style='text-align:center;vertical-align:middle'>".$exibirRegistros[2].'</td>';
            $html .= "<td style='text-align:center;vertical-align:middle'>".$exibirRegistros[3].'</td>';
            $html .= "<td style='text-align:center;vertical-align:middle'>".$exibirRegistros[4].'</td>';
			$html .= '</tr>';
			;
		}

		$sqlTotal = "
		SELECT 'Total', SUM(QtdImoveis) AS TotalQtdImoveis, CONCAT('R$ ',(REPLACE((SUM(Tarifa)), '.', ','))) AS TotalTarifa, CONCAT('R$ ',(REPLACE((SUM(Acrescimo)), '.', ','))) AS TotalAcrescimo, CONCAT('R$ ',(REPLACE((SUM(Total)), '.', ','))) AS TotalTotal  FROM 
		(SELECT Municipe.NomeMunicipe, COUNT(DISTINCT LoteCobrancaSerasa.CodImovel) AS QtdImoveis, SUM(Cobranca.VlrTarifa) AS Tarifa, SUM(Cobranca.VlrAcrescimo) AS Acrescimo, SUM(Cobranca.VlrTarifa + Cobranca.VlrAcrescimo) AS Total
		FROM LoteCobrancaSerasa 
		INNER JOIN CobrancaSerasa ON (LoteCobrancaSerasa.NrDocto = CobrancaSerasa.NrConv) 
		INNER JOIN Cobranca ON (CobrancaSerasa.NrCobrancaOrigem = Cobranca.NrCobranca) 
		INNER JOIN Municipe ON (LoteCobrancaSerasa.CodMunicipe = Municipe.CodMunicipe)
		WHERE CobrancaSerasa.Situacao = 'A'
		GROUP BY LoteCobrancaSerasa.CodMunicipe
		ORDER BY Municipe.NomeMunicipe
		) consulta;
		";
		$resultadoSqlTotal = mysqli_query($conexao6 , $sqlTotal);
		while($exibirRegistrosSqlTotal = mysqli_fetch_array($resultadoSqlTotal)){
			$html .= '<tr><strong>';
			$html .= "<td style='text-align:right;vertical-align:middle;font-weight: bold'>".$exibirRegistrosSqlTotal[0].'</td>';
			$html .= "<td style='text-align:center;vertical-align:middle;font-weight: bold'>".$exibirRegistrosSqlTotal[1].'</td>';
			$html .= "<td style='text-align:center;vertical-align:middle;font-weight: bold'>".$exibirRegistrosSqlTotal[2].'</td>';
            $html .= "<td style='text-align:center;vertical-align:middle;font-weight: bold'>".$exibirRegistrosSqlTotal[3].'</td>';
            $html .= "<td style='text-align:center;vertical-align:middle;font-weight: bold'>".$exibirRegistrosSqlTotal[4].'</td>';
			$html .= '</strong></tr>';
			;
		}

		header ("Expires: Mon, 26 Jul 1997 05:00:00 GMT");
		header ("Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT");
		header ("Cache-Control: no-cache, must-revalidate");
		header ("Pragma: no-cache");
		header ("Content-type: application/x-msexcel");
		header ("Content-Disposition: attachment; filename=\"{$arquivo}\"" );
		header ("Content-Description: PHP Generated Data" );

		echo $html;
		exit; ?>
	</body>
</html>