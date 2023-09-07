<?php

declare(strict_types=1);

/**
 * @var \Fr3nch13\Excel\View\ExcelBaseView $this
 */

$this->extend('/base');

if (!$this->fetch('page-title')) {
    $this->assign('page-title', __('Phpstans'));
}
if (!$this->fetch('page-subtitle')) {
    $this->assign('page-subtitle', __('All'));
}

$headers = ['ID', 'Name', 'Desc'];
$rows = [];
for ($i = 0; $i++; $i < 10) {
    $rows[$i] = [
        $i,
        __('Name {0}', [$i]),
        __('Desc {0}', [$i]),
    ];
}

$this->start('page-content');
?>
<section class="page-content">
    <div class="box">
        <!-- /.box-header -->
        <div class="box-body">
            <?php
            ?>
            <table>
                <?php
                echo $this->Html->tableHeaders($headers);
                echo $this->Html->tableCells($rows);
                ?>
            </table>
        </div>
    </div>
</section>
<?php $this->end(); /* page-content */
