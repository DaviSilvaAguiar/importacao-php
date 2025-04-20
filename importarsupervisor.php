<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

function limpar($v) {
    return trim(str_replace("'", "", $v ?? ""));
}

function formatarTelefone($v) {
    return preg_replace('/\D/', '', $v ?? "");
}

function formatarCNPJCPF($v) {
    return preg_replace('/\D/', '', $v ?? "");
}

function buscarTipo($v) {
    return strlen($v) === 11 ? 'F' : 'J';
}

function formatarData($v) {
    if (!$v) return null;
    $data = \DateTime::createFromFormat('d/m/Y', $v);
    return $data ? $data->format('Y-m-d') : null;
}

function normalizarCidade($v) {
    $v = mb_strtolower($v, 'UTF-8');
    $v = str_replace(
        ['á', 'é', 'í', 'ó', 'ú', 'ã', 'õ', 'â', 'ê', 'ô', 'ç'],
        ['a', 'e', 'i', 'o', 'u', 'a', 'o', 'a', 'e', 'o', 'c'],
        $v
    );
    $v = str_replace(["’", "'", "`", "´"], "", $v); // remove apóstrofos
    $v = preg_replace('/[^a-z\s]/', '', $v); // remove símbolos e números
    $v = preg_replace('/\s+/', ' ', $v); // remove espaços duplicados
    $v = trim($v);
    return ucwords($v); // tipo: "Herval D Oeste"
}

$transacaoIniciada = false;

try {
    // 1) Carregar Excel
    $arquivo = 'SUPERVISORES.xlsx';
    $spreadsheet = IOFactory::load($arquivo);
    $sheet = $spreadsheet->getActiveSheet();
    $dados = $sheet->toArray();

    // 2) Mapear colunas
    $cabecalho = $dados[0];
    $campos = [];
    foreach ($cabecalho as $i => $col) {
        $campos[trim($col)] = $i;
    }

    // 3) Conexão com o banco
    $pdo = new PDO("mysql:host=localhost;dbname=exemplo_db", "root", "");
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $pdo->beginTransaction();
    $transacaoIniciada = true;

    // 4) Zerar tabela
    $pdo->exec("SET FOREIGN_KEY_CHECKS = 0");
    $pdo->exec("TRUNCATE TABLE cad_supervisor");
    $pdo->exec("SET FOREIGN_KEY_CHECKS = 1");

    // 5) Loop de importação
    for ($i = 1; $i < count($dados); $i++) {
        $linha = $dados[$i];

        $nome = limpar($linha[$campos["NOME COMPLETO:"]]);
        if (empty($nome)) {
            continue; // Pula linhas sem nome
        }

        // Obter o valor bruto do campo e extrair o nome da empresa
        $idEmpresaRaw = limpar($linha[$campos["EMPRESA:"]]);
        $empresaParts = explode('-', $idEmpresaRaw);
        $nomeEmpresa = count($empresaParts) > 1 ? trim($empresaParts[1]) : trim($idEmpresaRaw);

        // Consultar na tabela cad_empresa para obter o id real da empresa
        $stmtEmpresa = $pdo->prepare("SELECT id_empresa FROM cad_empresa WHERE nome_empresa = :nome LIMIT 1");
        $stmtEmpresa->execute([':nome' => $nomeEmpresa]);
        $empresa = $stmtEmpresa->fetch(PDO::FETCH_ASSOC);
        if (!$empresa) {
            // Se a empresa não for encontrada, seta o id 33 como padrão para a fk_id_empresa
            $idEmpresaReal = 33;
        } else {
            $idEmpresaReal = $empresa['id_empresa'];
        }

        $areaFormacao = limpar($linha[$campos["ÁREA DE FORMAÇÃO:"]]);
        $tempoExperiencia = limpar($linha[$campos['TEMPO DE EXPERIÊNCIA:']]);

        $sql = "INSERT INTO cad_supervisor (
            nome_supervisor, fk_id_empresa, area_formacao, tempo_experiencia
        ) VALUES (
            :nome, :id_empresa, :area_formacao, :tempo_experiencia
        )";
        $stmt = $pdo->prepare($sql);
        $stmt->execute([
            ':nome' => $nome,
            ':id_empresa' => $idEmpresaReal,
            ':area_formacao' => $areaFormacao,
            ':tempo_experiencia' => $tempoExperiencia
        ]);

        echo "✅ Linha $i importada com sucesso\n";
    }

    $pdo->commit();
    echo "\n🎉 Importação concluída com sucesso!\n";

} catch (Exception $e) {
    echo "❌ Erro: " . $e->getMessage() . "\n";
    if (isset($pdo) && $transacaoIniciada) {
        $pdo->rollBack();
    }
}
