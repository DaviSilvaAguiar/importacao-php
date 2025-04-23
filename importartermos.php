<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

function limpar($v) {
    $v = $v ?? "";
    $v = trim(str_replace("'", "", $v));
    return $v === "" ? 0 : $v;
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
    return ucwords(trim($v));
}

function extrairNome($campo) {
    $partes = explode(' - ', $campo);
    return isset($partes[1]) ? trim($partes[1]) : trim($campo);
}

function buscarPorAproximacao(PDO $pdo, string $tabela, string $campoId, string $campoNome, string $valorOriginal) {
    $valor = trim($valorOriginal);
    $fragmentos = explode(' ', $valor);
    $busca = '';

    foreach ($fragmentos as $i => $palavra) {
        $busca .= ($i > 0 ? ' ' : '') . $palavra;
        $stmt = $pdo->prepare("SELECT {$campoId} FROM {$tabela} WHERE {$campoNome} LIKE :busca LIMIT 1");
        $stmt->execute([':busca' => $busca . '%']);
        $resultado = $stmt->fetch(PDO::FETCH_ASSOC);

        if ($resultado) {
            return $resultado[$campoId];
        }
    }

    return null;
}

$transacaoIniciada = false;

try {
    $arquivo = 'TERMOS.xlsx';
    $spreadsheet = IOFactory::load($arquivo);
    $sheet = $spreadsheet->getActiveSheet();
    $dados = $sheet->toArray();

    $cabecalho = $dados[0];
    $campos = [];
    foreach ($cabecalho as $i => $col) {
        $campos[trim($col)] = $i;
    }

    $pdo = new PDO("mysql:host=localhost;dbname=exemplo_db", "root", "");
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $pdo->beginTransaction();
    $transacaoIniciada = true;

    $pdo->exec("SET FOREIGN_KEY_CHECKS = 0");
    $pdo->exec("TRUNCATE TABLE cad_termos");
    $pdo->exec("SET FOREIGN_KEY_CHECKS = 1");

    for ($i = 1; $i < count($dados); $i++) {
        $linha = $dados[$i];

        // Estagiário
        $nomeEstagiario = extrairNome(limpar($linha[$campos["Nome do Estagiário"]]));
        if (empty($nomeEstagiario)) {
            echo "⚠️ Estagiário não informado (Linha $i). Linha ignorada.\n";
            continue;
        }

        $idEstagiarioReal = buscarPorAproximacao($pdo, 'cad_estagiarios', 'id_estagiario', 'nome_estagiario', $nomeEstagiario);
        if (!$idEstagiarioReal) {
            echo "⚠️ Estagiário não encontrado: $nomeEstagiario (Linha $i). Linha ignorada.\n";
            continue;
        }

        // Empresa
        $idEmpresaRaw = limpar($linha[$campos["Selecione a Empresa:"]]);
        $nomeEmpresa = extrairNome($idEmpresaRaw);
        $stmtEmpresa = $pdo->prepare("SELECT id_empresa FROM cad_empresas WHERE nome_empresa = :nome LIMIT 1");
        $stmtEmpresa->execute([':nome' => $nomeEmpresa]);
        $empresa = $stmtEmpresa->fetch(PDO::FETCH_ASSOC);
        $idEmpresaReal = $empresa['id_empresa'] ?? 33;

        // Escola
        $idEscola = extrairNome(limpar($linha[$campos["Selecione a Instituição de Ensino:"]]));
        $idEscolaReal = buscarPorAproximacao($pdo, 'cad_escolas', 'id_escola', 'nome_escola', $idEscola);
        if (!$idEscolaReal) {
            echo "⚠️ Escola não encontrada: $idEscola (Linha $i). Linha ignorada.\n";
            continue;
        }

        // Supervisor
        $idSupervisor = extrairNome(limpar($linha[$campos['Selecione o Supervisor(a):']]));
        $idSupervisorReal = buscarPorAproximacao($pdo, 'cad_supervisores', 'id_supervisor', 'nome_supervisor', $idSupervisor);
        if (!$idSupervisorReal) {
            echo "⚠️ Supervisor não encontrado: $idSupervisor (Linha $i). Linha ignorada.\n";
            continue;
        }

        // Outros campos
        $nomeOrientador = limpar($linha[$campos['Nome do Orientador']]);
        $cargoOrientador = limpar($linha[$campos['Cargo do Orientador']]);
        $dataInicio = formatarData(limpar($linha[$campos['Data de início do estágio:']]));        
        $dataFim = formatarData(limpar($linha[$campos['Data final do estágio:']]));        
        if (empty($dataInicio) || $dataInicio === '0000-00-00') {
            echo "⚠️ Linha $i ignorada: Data de início inválida ou ausente.\n";
            continue;
        }

        $valorBolsa = limpar($linha[$campos['Valor da Bolsa:']]);
        $valorBolsa = is_numeric($valorBolsa) ? $valorBolsa : 0;

        $sql = "INSERT INTO cad_termos (
            fk_id_estagiario, fk_id_empresa, fk_id_escola, fk_id_supervisor,
            nome_orientador, cargo_orientador, data_inicio_estagio, data_fim_estagio,
            valor_bolsa
        ) VALUES (
            :id_estagiario, :id_empresa, :id_escola, :id_supervisor,
            :nome_orientador, :cargo_orientador, :data_inicio_estagio, :data_fim_estagio,
            :valor_bolsa
        )";
        $stmt = $pdo->prepare($sql);
        $stmt->execute([
            ':id_estagiario'    => $idEstagiarioReal,
            ':id_empresa'       => $idEmpresaReal,
            ':id_escola'        => $idEscolaReal,
            ':id_supervisor'    => $idSupervisorReal,
            ':nome_orientador'  => $nomeOrientador,
            ':cargo_orientador' => $cargoOrientador,
            ':data_inicio_estagio' => $dataInicio,
            ':data_fim_estagio' => $dataFim,
            ':valor_bolsa'      => $valorBolsa
        ]);

        echo "✅ Linha $i importada com sucesso\n";
    }

    $pdo->commit();
    echo "\n🎉 Importação concluída com sucesso!\n";

} catch (Exception $e) {
    echo "❌ Erro: " . $e->getMessage() . "\n";
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
}
