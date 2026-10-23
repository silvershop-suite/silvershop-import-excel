<?php

declare(strict_types=1);

namespace SilverShop\Import\Excel\Task;

use SilverShop\Import\Excel\Export\ExcelCatalogueExporter;
use SilverShop\Import\Excel\Export\ExcelResultWriter;
use SilverShop\Import\Model\ImportProfile;
use SilverStripe\Control\Director;
use SilverStripe\Dev\BuildTask;
use SilverStripe\PolyExecution\PolyOutput;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;

/**
 * Produces a workbook — the export side of the roundtrip. {@code --type=result} writes a profile's staged
 * rows with status/error columns (fix + re-upload); {@code --type=catalogue} writes the current catalogue
 * as an editing template. Defaults to a timestamped file under {@code import-exports/} when no path given.
 */
class ExcelExportTask extends BuildTask
{
    protected static string $commandName = 'SilverShop-Import-Excel-Export';

    protected string $title = 'Import (Excel): export a result or catalogue workbook';

    protected static string $description =
        'Writes an .xlsx. --type=result --profile=<id|key> for the fix-and-re-upload workbook; '
        . '--type=catalogue for a full export. --domain=<name> (default Product). --path=<file> optional.';

    public function getOptions(): array
    {
        return [
            new InputOption('type', null, InputOption::VALUE_REQUIRED, 'result | catalogue (default: catalogue).'),
            new InputOption('profile', null, InputOption::VALUE_REQUIRED, 'Profile ID/ProfileKey (for --type=result).'),
            new InputOption('domain', null, InputOption::VALUE_REQUIRED, 'Domain to export (default: Product).'),
            new InputOption('path', null, InputOption::VALUE_REQUIRED, 'Output .xlsx path (default: auto-named).'),
        ];
    }

    protected function execute(InputInterface $input, PolyOutput $output): int
    {
        $type = (string) ($input->getOption('type') ?? 'catalogue') ?: 'catalogue';
        $domain = (string) ($input->getOption('domain') ?? 'Product') ?: 'Product';
        $path = (string) ($input->getOption('path') ?? '');
        if ($path === '') {
            $path = $this->defaultPath($type);
        }
        $this->ensureDir($path);

        if ($type === 'result') {
            $profile = $this->resolveProfile((string) ($input->getOption('profile') ?? ''));
            if (!$profile instanceof ImportProfile) {
                $output->writeln('--type=result needs a valid --profile. ' . $this->profileList());

                return Command::FAILURE;
            }
            $count = ExcelResultWriter::create()->write($profile, $domain, $path);
            $output->writeln(sprintf('Wrote %d staged row(s) for "%s" → %s', $count, (string) $profile->Title, $path));

            return Command::SUCCESS;
        }

        $count = ExcelCatalogueExporter::create()->export($path, $domain);
        $output->writeln(sprintf('Exported %d %s record(s) → %s', $count, $domain, $path));

        return Command::SUCCESS;
    }

    private function resolveProfile(string $ref): ?ImportProfile
    {
        if ($ref === '') {
            return null;
        }
        if (ctype_digit($ref)) {
            $byId = ImportProfile::get()->byID((int) $ref);
            if ($byId instanceof ImportProfile) {
                return $byId;
            }
        }

        return ImportProfile::get()->filter('ProfileKey', $ref)->first();
    }

    private function defaultPath(string $type): string
    {
        $base = Director::publicFolder() ?: Director::baseFolder();

        return rtrim($base, '/') . '/import-exports/' . $type . '-' . date('Ymd-His') . '.xlsx';
    }

    private function ensureDir(string $path): void
    {
        $dir = dirname($path);
        if (!is_dir($dir)) {
            mkdir($dir, 0775, true);
        }
    }

    private function profileList(): string
    {
        $rows = [];
        foreach (ImportProfile::get() as $profile) {
            $rows[] = sprintf('%d:%s', (int) $profile->ID, (string) ($profile->ProfileKey ?: $profile->Title));
        }

        return 'Available: ' . ($rows === [] ? '(none)' : implode(', ', $rows));
    }
}
