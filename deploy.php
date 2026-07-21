<?php
/**
 * deploy.php - Script de deploy automatico do data-entradas.js
 * Instalacao unica: coloque este arquivo em /compras/analise/deploy.php
 * 
 * Seguranca: protegido por token secreto
 */

$SECRET_TOKEN = "BrasildosParafusos2026!deploy";
$TARGET_FILE  = __DIR__ . '/data-entradas.js';

// Verificar token
$token = $_POST['token'] ?? $_GET['token'] ?? '';
if ($token !== $SECRET_TOKEN) {
    http_response_code(403);
    echo json_encode(['success' => false, 'error' => 'Token invalido']);
    exit;
}

// Upload do arquivo
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['data_file'])) {
    $uploaded = $_FILES['data_file'];
    
    if ($uploaded['error'] !== UPLOAD_ERR_OK) {
        echo json_encode(['success' => false, 'error' => 'Erro no upload: ' . $uploaded['error']]);
        exit;
    }
    
    if (!str_ends_with($uploaded['name'], '.js')) {
        echo json_encode(['success' => false, 'error' => 'Apenas arquivos .js sao aceitos']);
        exit;
    }
    
    if (move_uploaded_file($uploaded['tmp_name'], $TARGET_FILE)) {
        $size = round(filesize($TARGET_FILE) / 1024);
        echo json_encode([
            'success' => true, 
            'message' => "data-entradas.js atualizado! ({$size} KB)",
            'timestamp' => date('Y-m-d H:i:s')
        ]);
    } else {
        echo json_encode(['success' => false, 'error' => 'Falha ao mover arquivo para ' . $TARGET_FILE]);
    }
    exit;
}

// GET: status do arquivo atual
if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if (file_exists($TARGET_FILE)) {
        $size = round(filesize($TARGET_FILE) / 1024);
        $modified = date('Y-m-d H:i:s', filemtime($TARGET_FILE));
        echo json_encode([
            'success' => true,
            'file' => 'data-entradas.js',
            'size_kb' => $size,
            'last_modified' => $modified
        ]);
    } else {
        echo json_encode(['success' => false, 'error' => 'data-entradas.js nao encontrado']);
    }
    exit;
}

echo json_encode(['success' => false, 'error' => 'Metodo nao suportado']);