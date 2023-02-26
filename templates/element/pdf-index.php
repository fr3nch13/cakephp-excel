<?php

declare(strict_types=1);

/**
 * @var \Fr3nch13\Excel\View\PdfView $this
 */

$properties = $properties ?? [];
$headers = $headers ?? [];
$rows = $rows ?? [];

$this->Excel->create($properties, $headers, $rows)
    ->download('Pdf');
