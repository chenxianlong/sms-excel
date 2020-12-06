<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use AlibabaCloud\Client\AlibabaCloud;
use AlibabaCloud\Client\Exception\ClientException;
use AlibabaCloud\Client\Exception\ServerException;
class excel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'ccdgut:sendSMS {filePath}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'ccdgut:sendSMS';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        AlibabaCloud::accessKeyClient(env('ALIBABA_CLOUD_ACCESS_KEY_ID'), env('ALIBABA_CLOUD_ACCESS_SECRET'))
                        ->regionId('cn-hangzhou')
                        ->asDefaultClient();
        $reader = ReaderEntityFactory::createReaderFromFile($filePath = $this->argument("filePath"));
        $reader->open($filePath);
        foreach ($reader->getSheetIterator() as $sheet)
        {
            foreach ($sheet->getRowIterator() as $row)
            {
                // do stuff with the row
                $cells = $row->getCells();
                while(true)
                {
                    $name = $cells[2];
                    $phone = $cells[3];
                    try {
                        $result = AlibabaCloud::rpc()
                                            ->product('Dysmsapi')
                                            // ->scheme('https') // https | http
                                            ->version('2017-05-25')
                                            ->action('SendSms')
                                            ->method('POST')
                                            ->host('dysmsapi.aliyuncs.com')
                                            ->options([
                                                            'query' => [
                                                            'RegionId' => "cn-hangzhou",
                                                            'PhoneNumbers' => $phone,
                                                            'SignName' => "东莞理工学院城市学院",
                                                            'TemplateCode' => "SMS_190276251",
                                                            'TemplateParam' => "{\"name\":\"{$name}\"}",
                                                            ],
                                                        ])
                                            ->request();
                        print_r($result->toArray());
                    } catch (ClientException $e) {
                        echo $e->getErrorMessage() . PHP_EOL;
                    } catch (ServerException $e) {
                        echo $e->getErrorMessage() . PHP_EOL;
                    }
                    break;
                }
            }
        }
        $reader->close();
    }
}
