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
    $v = str_replace(["’", "'", "`", "´"], "", $v);
    $v = preg_replace('/[^a-z\s]/', '', $v);
    $v = preg_replace('/\s+/', ' ', $v);
    $v = trim($v);
    return ucwords($v);
}

$transacaoIniciada = false;

try {
    $arquivo = 'EMPRESAS.xlsx';
    $spreadsheet = IOFactory::load($arquivo);
    $sheet = $spreadsheet->getActiveSheet();
    $dados = $sheet->toArray();

    $cabecalho = $dados[0];
    foreach ($cabecalho as $i => $col) {
        $campos[trim($col)] = $i;
    }

    $pdo = new PDO("mysql:host=localhost;dbname=exemplo_db", "root", "");
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $pdo->beginTransaction();
    $transacaoIniciada = true;

    $pdo->exec("SET FOREIGN_KEY_CHECKS = 0");
    $pdo->exec("TRUNCATE TABLE cad_empresa");
    $pdo->exec("SET FOREIGN_KEY_CHECKS = 1");

    for ($i = 1; $i < count($dados); $i++) {
        $linha = $dados[$i];

        // Ignorar linhas completamente vazias
        if (empty(array_filter($linha))) {
            continue;
        }

        $nome = limpar($linha[$campos["NOME DA EMPRESA CONCEDENTE:"]]);
        $cnpj = formatarCNPJCPF($linha[$campos["CNPJ"]]);

        // Ignorar se não houver nome ou CNPJ
        if (empty($nome) || empty($cnpj)) {
            echo "⚠️ Linha $i ignorada (sem nome ou CNPJ)\n";
            continue;
        }

        $telefone = formatarTelefone($linha[$campos["NÚMERO DE CONTATO"]]);
        $cep = formatarCNPJCPF($linha[$campos["CEP"]]);
        $cep = (strlen($cep) > 9 || empty($cep)) ? "00000-000" : $cep;

        $logradouro = limpar($linha[$campos["LOGRADOURO"]]);
        $numero = limpar($linha[$campos['NÚMERO']]);
        $numero = empty($numero) ? 'S/N' : $numero;

        $complemento = limpar($linha[$campos["COMPLEMENTO:"]]);
        $bairro = limpar($linha[$campos["BAIRRO:"]]);
        $cidadeNome = limpar($linha[$campos["CIDADE:"]]);

        $nomeRepresentante = limpar($linha[$campos["NOME COMPLETO:"]]);
        $cargoRepresentante = limpar($linha[$campos["CARGO:"]]);
        $cpfRepresentante = formatarCNPJCPF($linha[$campos["CPF"]]);

        $idCidade = 4588; // Padrão: Tubarão

        if (!empty($cidadeNome)) {
            $cidadeNomeNormalizada = normalizarCidade($cidadeNome);
            if ($cidadeNomeNormalizada == "Herval D Oeste") {
                $idCidade = 4416;
            } else {
                $stmtCidade = $pdo->prepare("SELECT id_cidade FROM tb_cidade WHERE REPLACE(REPLACE(LOWER(nm_cidade), '’', ''), '''', '') = :nome LIMIT 1");
                $stmtCidade->execute([':nome' => strtolower(str_replace(["’", "'"], '', $cidadeNomeNormalizada))]);
                $cidadeRow = $stmtCidade->fetch(PDO::FETCH_ASSOC);
                if ($cidadeRow) {
                    $idCidade = $cidadeRow['id_cidade'];
                } else {
                    echo "⚠️ Cidade '{$cidadeNome}' não encontrada. Usando padrão Tubarão (4588)\n";
                }
            }
        }

        $sql = "INSERT INTO cad_empresa (
            nome_empresa, numero_cnpj, numero_telefone,
            numero_cep, endereco, numero_endereco, complemento_endereco,
            bairro, fk_id_cidade, nome_representante, cargo_representante,
            cpf_representante
        ) VALUES (
            :nome, :cnpj, :telefone,  
            :cep, :endereco, :numero, :complemento, 
            :bairro, :cidade, :nome_representante, :cargo_representante,
            :cpf_representante
        )";

        $stmt = $pdo->prepare($sql);
        $stmt->execute([
            ':nome' => $nome,
            ':cnpj' => $cnpj,
            ':telefone' => $telefone,
            ':cep' => $cep, 
            ':endereco' => $logradouro,
            ':numero' => $numero,
            ':complemento' => $complemento,
            ':bairro' => $bairro,
            ':cidade' => $idCidade,
            ':nome_representante' => $nomeRepresentante,
            ':cargo_representante' => $cargoRepresentante,  
            ':cpf_representante' => $cpfRepresentante
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
