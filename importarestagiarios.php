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
    $arquivo = 'ESTAGIARIOS.xlsx';
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
    $pdo->exec("TRUNCATE TABLE cad_estagiario");
    $pdo->exec("SET FOREIGN_KEY_CHECKS = 1");

    // 5) Loop de importação
    for ($i = 1; $i < count($dados); $i++) {
        $linha = $dados[$i];

        $nome = limpar($linha[$campos["NOME COMPLETO:"]]);
        $cpf = formatarCNPJCPF($linha[$campos["CPF"]]);
        $dataNascimento = formatarData($linha[$campos["DATA DE NASCIMENTO:"]]);
        $telefone = formatarTelefone($linha[$campos["NÚMERO DE TELEFONE/CELULAR"]]);
        $email = limpar($linha[$campos["Endereço de e-mail"]]);
        $cep = formatarCNPJCPF($linha[$campos["CEP"]]);
        if (strlen($cep) > 9 || empty($cep)) {
            $cep = "00000-000"; // Valor padrão para CEP inválido
        }
        $logradouro = limpar($linha[$campos["NOME DO LOGRADOURO"]]);
        $numero = limpar($linha[$campos['NÚMERO']]);
        if (empty($numero)) {
            $numero = 'S/N'; // Valor padrão para número do endereço
        }
        $complemento = limpar($linha[$campos["COMPLEMENTO:"]]);
        $bairro = limpar($linha[$campos["BAIRRO:"]]);
        $cidadeNome = limpar($linha[$campos["CIDADE:"]]);
        $curso = limpar($linha[$campos["CURSO:"]]);
        $nivelCurso = limpar($linha[$campos["NÍVEL DO CURSO:"]]);

        // Verificar e ajustar o nível do curso
        if (strpos($nivelCurso, "Pós-graduação") !== false) {
            $nivelCurso = "Pós-graduação"; // Ajustar para evitar erro de tamanho
        }

        $setorEstagio = limpar($linha[$campos["Em que setor você está estagiando ou pretende estagiar?"]]);
        $nomeMae = limpar($linha[$campos["NOME COMPLETO DA MÃE:"]]);
        $comprovanteResidencia = limpar($linha[$campos["COMPROVANTE DE RESIDÊNCIA"]]);
        $docEscolar = limpar($linha[$campos["DOCUMENTO ESCOLAR/UNIVERSIDADE"]]);
        $pis = limpar($linha[$campos["NUMERO DO PIS"]]);
        $chavePix = limpar($linha[$campos["CHAVE PIX"]]);
        $instituicao = limpar($linha[$campos["INSTITUIÇÃO DE ENSINO:"]]);

        // 🔍 Verificar cidade (padrão: Tubarão - ID 4588)
        $idCidade = 4588;

        if (!empty($cidadeNome)) {
            $cidadeNomeNormalizada = normalizarCidade($cidadeNome);
            if ($cidadeNomeNormalizada == "Herval D Oeste") {
                $idCidade = 4416;  // ID de "Herval D'Oeste"
            } else {
                $stmtCidade = $pdo->prepare("SELECT id_cidade FROM cad_cidade WHERE REPLACE(REPLACE(LOWER(nm_cidade), '’', ''), '''', '') = :nome LIMIT 1");
                $stmtCidade->execute([':nome' => strtolower(str_replace(["’", "'"], '', $cidadeNomeNormalizada))]);
                $cidadeRow = $stmtCidade->fetch(PDO::FETCH_ASSOC);

                if ($cidadeRow) {
                    $idCidade = $cidadeRow['id_cidade'];
                } else {
                    echo "⚠️ Cidade '{$cidadeNome}' não encontrada. Usando padrão Tubarão (4588)\n";
                }
            }
        }

        // Inserir estagiário
        $sql = "INSERT INTO cad_estagiario (
            nome_estagiario, numero_cpf, data_nascimento, numero_telefone, numero_celular,
            email, numero_cep, endereco, numero_endereco, complemento_endereco,
            bairro, fk_id_cidade, curso, nivel_curso, area_de_estagio, nome_mae, foto_documento, 
            comprovante_residencia, numero_pis, chave_pix, instituicao_ensino
        ) VALUES (
            :nome, :cpf, :data_nascimento, :telefone, :telefone, 
            :email, :cep, :endereco, :numero, :complemento, 
            :bairro, :cidade, :curso, :nivel, :area, :mae, :documento, 
            :comprovante, :pis, :pix, :instituicao
        )";

        $stmt = $pdo->prepare($sql);
        $stmt->execute([
            ':nome' => $nome,
            ':cpf' => $cpf,
            ':data_nascimento' => $dataNascimento,
            ':telefone' => $telefone,
            ':email' => $email,
            ':cep' => $cep, // CEP corrigido
            ':endereco' => $logradouro,
            ':numero' => $numero, // Número corrigido
            ':complemento' => $complemento,
            ':bairro' => $bairro,
            ':cidade' => $idCidade,
            ':curso' => $curso,
            ':nivel' => $nivelCurso,
            ':area' => $setorEstagio,
            ':mae' => $nomeMae,
            ':documento' => $docEscolar,
            ':comprovante' => $comprovanteResidencia,
            ':pis' => $pis,
            ':pix' => $chavePix,
            ':instituicao' => $instituicao
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
