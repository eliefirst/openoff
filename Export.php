<?php
/**
 * This file is part of Rl_Sales for Magento.
 *
 * @license   Tous droits réservés
 * @author    Elie Laurent (elie@redline.paris)
 * @category  Rl
 * @package   Rl_Sales
 * @copyright Copyright (c) 2021 Elielweb
 */


namespace Rl\Sales\Model\Order\Odeis;

use Clrz\Toolbox\Helper\Data as ClrzToolbox;
use Magento\Catalog\Api\ProductRepositoryInterface as ProductRepository;
use Magento\Catalog\Model\ResourceModel\Product as ProductResourceModel;
use Magento\Directory\Model\CountryFactory;
use Magento\Framework\App\Filesystem\DirectoryList;
use Magento\Framework\Exception\NoSuchEntityException;
use Magento\Framework\Filesystem\Io\File;
use Rl\Sales\Helper\Odeis as OdeisHelper;

class Export
{
    /** @var \Rl\Sales\Helper\Odeis $_odeisHelper */
    protected $_odeisHelper;
    /** @var \Clrz\Toolbox\Helper\Data $_clrzToolboxHelper */
    protected $_clrzToolboxHelper;
    /** @var \Magento\Framework\App\Filesystem\DirectoryList $_directoryList */
    protected $_directoryList;
    /** @var \Magento\Framework\Filesystem\Io\File $_file */
    protected $_file;
    /** @var string $_exportDirectory */
    protected $_exportDirectory;
    /** @var \Magento\Directory\Model\CountryFactory $_countryFactory */
    protected $_countryFactory;
    /** @var \Magento\Catalog\Api\ProductRepositoryInterface $_productRepository */
    protected $_productRepository;

    protected $productResourceModel;

    /**
     * Export constructor.
     *
     * @param \Rl\Sales\Helper\Odeis $odeisHelper
     * @param \Clrz\Toolbox\Helper\Data $clrzToolboxHelper
     * @param \Magento\Framework\Filesystem\Io\File $file
     * @param \Magento\Framework\App\Filesystem\DirectoryList $directoryList
     * @param \Magento\Directory\Model\CountryFactory $countryFactory
     * @param \Magento\Catalog\Api\ProductRepositoryInterface $productRepository
     * @param \Magento\Store\Model\App\Emulation $appEmulation
     */
    public function __construct(
        OdeisHelper $odeisHelper,
        ClrzToolbox $clrzToolboxHelper,
        File $file,
        DirectoryList $directoryList,
        CountryFactory $countryFactory,
        ProductRepository $productRepository,
        ProductResourceModel $productResourceModel
    ) {
        $this->_odeisHelper         = $odeisHelper;
        $this->_clrzToolboxHelper   = $clrzToolboxHelper;
        $this->_directoryList       = $directoryList;
        $this->_file                = $file;
        $this->_countryFactory      = $countryFactory;
        $this->_productRepository   = $productRepository;
        $this->productResourceModel = $productResourceModel;
    }

    public function getCountryName($countryCode)
    {
        $country = $this->_countryFactory->create()->loadByCode($countryCode);

        return $country->getName();
    }

    public function mapPaymentMethod($method)
    {
        $defaultMethod = 'CyberMUT-P@iement';

        $paymentMethodMapping = [
            'checkmo'          => 'Cheque',
            'monetico_onetime' => 'CyberMUT-P@iement',
            'banktransfer'     => 'Virement Bancaire',
            'adyen_cc'         => 'Adyen-P@iement',
            'adyen_hpp'        => 'Adyen-P@iement',
        ];

        return isset($paymentMethodMapping[ $method ]) ? $paymentMethodMapping[ $method ] : $defaultMethod;
    }

    public function getSizeList()
    {
        return [
            "15,5 cm",
            "15.5 cm",
            "16,5 cm",
            "16.5 cm",
            "17,5 cm",
            "17.5 cm",
            "18 cm",
            "18.0 cm Men",
            "18.5 cm",
            "20 cm",
            "20.0 cm Men",
            "23.5 cm",
            "24.5 cm",
            "25.5 cm",
            "30 cm 2 à 4 ans",
            "35 cm 5 à 12 ans",
            "38 Cm",
            "38 cm Ras du Cou",
            "39 Cm",
            "40 Cm",
            "42 cm",
            "43 cm",
            "43 cm Homme",
            "44 cm",
            "44 mm",
            "45 mm",
            "46 mm",
            "47 mm",
            "48 mm",
            "49 mm",
            "50 mm",
            "51 mm",
            "52 mm",
            "53 mm",
            "54 mm",
            "55 mm",
            "56 mm",
            "57 mm",
            "58 mm",
            "59 mm",
            "60 mm",
            "78 cm Sautoir",
            "78 cm Tour de Taille",
            "84 cm Tour de Taille",
            "Bébé 3 à 6 mois 10.5 cm",
            "Bébé 7 à 11 mois 11.5 cm",
            "Bébé 9 cm",
            "Enfant de 1 à 3 ans 12.5 cm",
            "Enfant de 10 à 14 ans 15.5 cm",
            "Enfant de 4 à 7 ans 13.5 cm",
            "Enfant de 8 à 12 ans 14.5 cm",
        ];
    }

    public function execute($orderCollection)
    {
        // Export directory
        $this->_exportDirectory = $this->_directoryList->getPath(
                DirectoryList::MEDIA
            ) . "/" . $this->_odeisHelper->getExportDirectory();

        // Create directory if not exists
        $this->_file->checkAndCreateFolder($this->_exportDirectory, 0700);

        // Create file
        $fileName = 'orders_export_' . date('Ymd_His') . '.cde';
        $filePath = $this->_exportDirectory . "/" . $fileName;

        // Init
        $content              = '';
        $eol                  = "\r\n";
        $toReplaceList        = ['链手镯', '链手镯和线', 'À', 'Á', 'Â', 'Ã', 'Ä', 'Å', 'Ç', 'È', 'É', 'Ê', 'Ë', 'Ì', 'Í', 'Î', 'Ï', 'Ò', 'Ó', 'Ô', 'Õ', 'Ö', 'Ù', 'Ú', 'Û', 'Ü', 'Ý', 'à', 'á', 'â', 'ã', 'ä', 'å', 'ç', 'è', 'é', 'ê', 'ë', 'ì', 'í', 'î', 'ï', 'ð', 'ò', 'ó', 'ô', 'õ', 'ö', 'ù', 'ú', 'û', 'ü', 'ý', 'ÿ', '项链', 'Rouge1', 'Licorice', 'CHAIN NECKLACE', '戒指', '项链纱', '链项链', 'チェーンの', 'ネックレス', 'Paiement par cheque', 'Mayssa', '  ', 'DOUBLE CHAIN AND STRING BRACELET', 'STRING BRACELET', '❤', '对唱', 'Œ', 'PURE DOUBLE ENFANT', 'Fedex-Express - Fedex-Express Without Insurance', 'Fedex-Express - Fedex-Priority Without Insurance', 'Pick Up Point - Colissimo Pick-Up-Point Insured', 'Pick Up Point - Colissimo Point de retrait Assuré', 'International - Colissimo Expert International Assuré', 'International - Colissimo Expert International Insured', 'Domicile - Colissimo Signature Insured', 'Domicile - Colissimo Signature Assuré', 'Colissimo -  National', 'Colissimo -  Colissimo OM1 Assure', 'Colissimo -  Colissimo OM2 Assure', 'Colissimo -  Colissimo I', 'Mon transporteur -  Mon transporteur', 'Livreur 75 -  Firstclass-Express(Coursier)', '宝宝 7至11个月 11.5厘米 cm', 'Fedex-Express -  ', '紫', '檀', '桃', '宝', '厘', '啊', '爱', '安', '暗', '按', '八', '把', '爸', '吧', '白', '百', '拜', '班', '般', '板', '半', '办', '帮', '包', '保', '環', '婴', 'ュ', '萤', '抱', '报', '爆', '杯', '北', '被', '背', '备', '本', '鼻', '蓝', '童', '糖', '蔚', '粉', '橙', '盆', '覆', '浆', '罂', '粟', '镀', '橄', '榄', '卡', '奶', '腳', '鍊', '葡', '萄', '比', '笔', '避', '必', '边', '便', '遍', '辨', '变', '标', '表', '别', '病', '并', '补', '不', '部', '布', '步', '才', '材', '采', '彩', '菜', '参', '草', '层', '曾', '茶', '察', '查', '差', '产', '长', '常', '场', '厂', '唱', '车', '彻', '称', '成', '城', '承', '程', '吃', '冲', '虫', '出', '初', '除', '楚', '处', '川', '穿', '传', '船', '窗', '床', '创', '春', '词', '此', '次', '从', '村', '存', '错', '答', '达', '打', '大', '带', '待', '代', '单', '但', '淡', '蛋', '当', '党', '导', '到', '道', '的', '得', '灯', '等', '低', '底', '地', '第', '弟', '点', '典', '电', '店', '掉', '调', '丁', '定', '冬', '东', '懂', '动', '都', '读', '独', '度', '短', '断', '段', '对', '队', '多', '朵', '躲', '饿', '儿', '而', '耳', '二', '发', '乏', '法', '反', '饭', '范', '方', '房', '防', '访', '放', '非', '飞', '费', '分', '坟', '份', '风', '封', '夫', '服', '福', '府', '父', '副', '复', '富', '妇', '该', '改', '概', '敢', '感', '干', '刚', '钢', '高', '搞', '告', '哥', '歌', '革', '隔', '格', '个', '给', '跟', '根', '更', '工', '公', '功', '共', '狗', '够', '构', '姑', '古', '骨', '故', '顾', '固', '瓜', '刮', '挂', '怪', '关', '观', '官', '馆', '管', '惯', '光', '广', '规', '鬼', '贵', '国', '果', '过', '还', '孩', '海', '害', '含', '汉', '好', '号', '喝', '河', '和', '何', '合', '黑', '很', '恨', '红', '后', '候', '呼', '忽', '乎', '湖', '胡', '虎', '户', '互', '护', '花', '华', '划', '画', '化', '话', '怀', '坏', '欢', '环', '换', '黄', '回', '会', '婚', '活', '火', '或', '货', '获', '机', '鸡', '积', '基', '极', '及', '集', '级', '急', '几', '己', '寄', '继', '际', '记', '济', '纪', '技', '计', '季', '家', '加', '假', '架', '价', '间', '简', '见', '建', '健', '件', '江', '将', '讲', '交', '饺', '脚', '角', '叫', '教', '较', '接', '街', '阶', '结', '节', '解', '姐', '介', '界', '今', '金', '斤', '仅', '紧', '近', '进', '尽', '京', '经', '精', '睛', '景', '静', '境', '究', '九', '酒', '久', '就', '旧', '救', '居', '局', '举', '句', '具', '据', '剧', '拒', '觉', '绝', '决', '军', '开', '看', '康', '考', '靠', '科', '可', '课', '刻', '客', '肯', '空', '孔', '口', '苦', '哭', '快', '筷', '块', '况', '困', '拉', '来', '浪', '劳', '老', '乐', '了', '累', '类', '冷', '离', '李', '里', '理', '礼', '立', '丽', '利', '历', '力', '例', '连', '联', '脸', '练', '凉', '两', '辆', '亮', '量', '谅', '疗', '料', '烈', '林', '零', '〇', '领', '另', '留', '流', '六', '龙', '楼', '路', '旅', '绿', '虑', '论', '落', '妈', '马', '吗', '买', '卖', '满', '慢', '忙', '毛', '么', '没', '美', '每', '门', '们', '猛', '梦', '迷', '米', '密', '面', '民', '名', '明', '命', '某', '母', '木', '目', '拿', '哪', '那', '男', '南', '难', '脑', '闹', '呢', '内', '能', '你', '年', '念', '娘', '鸟', '您', '牛', '农', '弄', '怒', '女', '暖', '怕', '排', '派', '判', '旁', '跑', '培', '朋', '皮', '篇', '片', '票', '品', '平', '评', '漂', '破', '普', '七', '期', '骑', '其', '奇', '齐', '起', '气', '汽', '器', '千', '前', '钱', '强', '墙', '桥', '巧', '切', '且', '亲', '轻', '青', '清', '情', '请', '庆', '穷', '秋', '求', '球', '区', '取', '去', '趣', '全', '缺', '却', '确', '然', '让', '扰', '热', '人', '认', '任', '日', '容', '肉', '如', '入', '三', '色', '杀', '山', '善', '商', '伤', '上', '少', '绍', '蛇', '设', '社', '谁', '身', '深', '什', '神', '甚', '生', '声', '升', '省', '师', '诗', '十', '时', '识', '实', '食', '始', '使', '史', '是', '事', '市', '室', '示', '似', '视', '适', '式', '士', '试', '世', '势', '收', '手', '守', '首', '受', '书', '舒', '熟', '数', '术', '树', '双', '水', '睡', '顺', '说', '思', '司', '私', '死', '四', '送', '诉', '算', '虽', '随', '岁', '碎', '所', '索', '他', '她', '它', '台', '太', '态', '谈', '特', '疼', '提', '题', '体', '替', '天', '田', '条', '铁', '听', '挺', '停', '通', '同', '统', '头', '突', '图', '土', '团', '推', '托', '外', '完', '玩', '晚', '碗', '万', '王', '往', '忘', '望', '为', '围', '委', '位', '卫', '味', '温', '文', '闻', '问', '我', '屋', '无', '五', '午', '武', '舞', '物', '务', '西', '息', '希', '析', '习', '喜', '洗', '细', '系', '下', '吓', '夏', '先', '鲜', '显', '现', '线', '限', '香', '乡', '相', '想', '响', '象', '向', '像', '项', '消', '小', '校', '笑', '效', '些', '鞋', '写', '谢', '新', '心', '信', '星', '行', '形', '醒', '姓', '兴', '幸', '性', '休', '修', '需', '许', '续', '选', '学', '雪', '血', '寻', '牙', '呀', '言', '研', '颜', '眼', '演', '验', '阳', '羊', '养', '样', '要', '药', '爷', '也', '夜', '叶', '业', '一', '医', '衣', '依', '疑', '以', '已', '意', '义', '艺', '忆', '易', '议', '因', '音', '阴', '印', '应', '英', '影', '硬', '映', '用', '优', '由', '油', '有', '友', '又', '右', '鱼', '于', '语', '雨', '与', '遇', '育', '欲', '元', '园', '原', '员', '园', '远', '院', '愿', '约', '月', '越', '云', '运', '杂', '在', '再', '咱', '早', '造', '则', '怎', '增', '展', '站', '张', '丈', '章', '招', '找', '照', '者', '这', '着', '真', '诊', '正', '整', '政', '证', '知', '之', '支', '织', '直', '职', '值', '只', '指', '纸', '止', '至', '制', '治', '致', '志', '中', '钟', '终', '种', '重', '众', '周', '洲', '州', '竹', '主', '住', '祝', '注', '著', '助', '专', '转', '庄', '装', '壮', '准', '资', '子', '仔', '字', '自', '总', '走', '租', '族', '足', '组', '嘴', '最', '昨', '左', '作', '做', '坐', '座', '纱', '荧', '戒', '洋', '列', '绳', '翡', '翠', '波', '罗', '樱', '苹', '珠', '兰', '巴', '黎', '灰', '勒', '链', 'あ', 'か', 'さ', 'た', 'な', 'は', 'ま', 'や', 'ら', 'わ', 'が', 'ざ', 'だ', 'ば', 'ぱ', 'ア', 'カ', 'サ', 'タ', 'ナ', 'ハ', 'マ', 'ヤ', 'ラ', 'ワ', 'ガ', 'ザ', 'ダ', 'バ', 'パ', 'い', 'き', 'し', 'ち', 'に', 'ひ', 'み', 'り', 'ゐ', 'ぎ', 'じ', 'ぢ', 'び', 'ぴ', 'イ', 'キ', 'シ', 'チ', 'ニ', 'ヒ', 'ミ', 'リ', 'ヰ', 'ギ', 'ジ', 'ヂ', 'ビ', 'ピ', 'う', 'く', 'す', 'つ', 'ぬ', 'ふ', 'む', 'ゆ', 'る', 'ぐ', 'ず', 'づ', 'ぶ', 'ぷ', 'ウ', 'ク', 'ス', 'ツ', 'ヌ', 'フ', 'ム', 'ユ', 'ル', 'グ', 'ズ', 'ヅ', 'ブ', 'プ', 'え', 'け', 'せ', 'て', 'ね', 'へ', 'め', 'れ', 'ゑ', 'げ', 'ぜ', '極', 'で', 'べ', 'ぺ', 'エ', 'ケ', 'セ', 'テ', 'ネ', 'ヘ', 'メ', 'レ', 'ヱ', 'ゲ', 'ゼ', 'デ', 'ベ', 'ペ', 'お', 'こ', 'そ', 'と', 'の', 'ほ', 'も', 'よ', 'ろ', 'を', 'ん', 'ご', 'ぞ', 'ど', 'ぼ', 'ぽ', '丝', 'オ', 'コ', 'ソ', 'ト', 'ノ', 'ホ', 'モ', 'ヨ', 'ロ', 'ヲ', 'ン', 'ゴ', 'ゾ', 'ド', 'ボ', 'ポ', 'ー', 'ッ', 'ェ', '蛍', '镯', '纹', '螺', ' FIL', 'Bracelet ', 'Enfant de', 'BRACELET', 'STRING ', '"', 'Mr. ', 'M. ', 'Mme ', 'Mrs. ', 'Mme. ', 'Mlle ', 'Fedex-Express -', 'CHAIN AND STRING '];
        $replacementList      = ['Bracelet Chaine', ' ', 'A', 'A', 'A', 'A', 'A', 'A', 'C', 'E', 'E', 'E', 'E', 'I', 'I', 'I', 'I', 'O', 'O', 'O', 'O', 'O', 'U', 'U', 'U', 'U', 'Y', 'a', 'a', 'a', 'a', 'a', 'a', 'c', 'e', 'e', 'e', 'e', 'i', 'i', 'i', 'i', 'o', 'o', 'o', 'o', 'o', 'o', 'u', 'u', 'u', 'u', 'y', 'y', 'collier', 'Rouge', 'Reglisse', 'Collier Chaine', 'Bague', 'Collier', 'Collier Chaine', 'Chaine', 'Collier', 'Cheque', 'Fleur', ' ', 'DOUBLE', 'FIL', 'COEUR', 'DUO', 'OE', 'DOUBLE PURE ENFANT', 'Fedex-Express', 'Fedex-Priority', 'Colissimo Point de retrait', 'Colissimo Point de retrait', 'Colissimo Expert International', 'Colissimo Expert International', 'Colissimo - National', 'Colissimo - National', 'Colissimo - National', 'Colissimo OM1 Assure', 'Colissimo OM2 Assure', 'Colissimo I', 'Mon transporteur', 'Firstclass-Express', 'BB 11.5cm', ''];
        $sizeList             = $this->getSizeList();
        $toReplaceCarrierList = ['Fedex - Express - Fedex - Express Without Insurance', 'Fedex - Express - Fedex - Express (sans assurance)', 'La Poste : Colissimo - Expert International (Avec Assurance)', 'La Poste : Colissimo - Points Relais (Avec Assurance)', 'La Poste : Colissimo - Domicile avec signature (Avec Assurance)', '郵便局：コリシモ - Expert International (Avec Assurance)', '邮：colissimo - Expert International (Avec Assurance)', 'Post office: Colissimo - Expert International (Avec Assurance)', '邮局：colissimo - Expert International (Avec Assurance)'];
        $carrierList          = ['Fedex-Express', 'Fedex-Express', 'Colissimo - National', 'Colissimo - National', 'Colissimo - National', 'Colissimo Expert International', 'Colissimo Expert International', 'Colissimo Expert International', 'Colissimo Expert International'];
        $toReplaceb2b         = ['checkmo'];
        $replacmentb2b        = ['0'];
        /* @var \Mage\Sales\Model\Order $order */
        foreach ($orderCollection as $order) {
            $items      = $order->getAllVisibleItems();
            $isB2cOrder = $this->_odeisHelper->isB2cOrder($order);
            $customerId = $order->getCustomerId();
            $orderDate  = new \DateTime($order->getCreatedAt());

            // ADDRESSES
            $billingAddress  = $order->getBillingAddress();
            $shippingAddress = $order->getShippingAddress();

            // HEADER
            $content .= '[ENTETE]' . $eol;
            $content .= 'TYPE = COMMANDE' . $eol;
            $acpt    = $order->getGrandTotal();
            //if  {
            //     $content .= 'ACPTE = ' . (
            //            str_replace ($toReplaceb2b, $replacmentb2b, $this->mapPaymentMethod($order->getPayment()->getMethod() . $eol;
            //    ) . $eol;
            // } else {

            //    $content .= 'ACPTE = ' . $acpt . $eol;
            //         ) . $eol;
            // }

            $content .= 'ACPTE = ' . $acpt . $eol;     //$order->getGrandTotal() . $eol;

            $content .= 'MODEREGL = ' . $this->mapPaymentMethod($order->getPayment()->getMethod()) . $eol;
            $content .= 'MODELIVR = ' . str_replace(
                    $toReplaceCarrierList, $carrierList, $order->getShippingDescription()
                ) . $eol;
            $content .= 'FRAIST = ' . $order->getShippingAmount() . $eol;
            $content .= 'EMAILFOUR = redline@edi.com' . $eol;
            $content .= 'CL CDE = W' . $customerId . $eol;
            $content .= 'SUFFIX = ' . ($order->getData('customer_prefix')) . $eol;
            $content .= 'DAT CDE = ' . $orderDate->format('dmY') . $eol;

            // Date livraison b2b / b2c
            if (!$isB2cOrder) {
                $content .= 'DAT LIV = ' . date('dmY', strtotime("+35 days")) . $eol;
            }
            else {
                $content .= 'DAT LIV = ' . date('dmY', strtotime("+10 days")) . $eol;
            }

            $content .= 'EMAIL = ' . $order->getCustomerEmail() . $eol;
            $content .= 'NO CDE = ' . $order->getIncrementId() . $eol;
            $content .= 'TYP CDE = ' . ($isB2cOrder ? "BTOC" : "BTOB") . $eol;
            $content .= 'DEVISE = ' . $order->getOrderCurrencyCode() . $eol;
            $content .= 'NBLIG = ' . $order->getData('total_item_count') . $eol;

            // BILLING ADDRESS
            if (!$isB2cOrder) {
                $content .= 'ONF_NOM = ' . strtoupper(
                        str_replace($toReplaceList, $replacementList, $billingAddress->getCompany())
                    ) . $eol;
            }
            else {
                $content .= 'ONF_NOM = ' . strtoupper(
                        str_replace($toReplaceList, $replacementList, $billingAddress->getName())
                    ) . $eol;
            }

            $streetArray = $billingAddress->getStreet();
            $address1    = isset($streetArray[0]) ? $streetArray[0] : "";
            $address2    = isset($streetArray[1]) ? $streetArray[1] : "";
            $city        = $billingAddress->getCity();
            $state       = $billingAddress->getRegion();
            $country     = $this->getCountryName($billingAddress->getCountryId());

            $content .= 'ONF_ADR1 = ' . str_replace($toReplaceList, $replacementList, $address1) . $eol;
            $content .= 'ONF_ADR2 = ' . str_replace($toReplaceList, $replacementList, $address2) . $eol;
            $content .= 'ONF_CP = ' . $billingAddress->getPostcode() . $eol;
            $content .= 'ONF_VILLE = ' . strtoupper(str_replace($toReplaceList, $replacementList, $city)) . $eol;
            $content .= 'ONF_ETAT = ' . str_replace($toReplaceList, $replacementList, $state) . $eol;
            $content .= 'ONF_PAYS = ' . strtoupper(str_replace($toReplaceList, $replacementList, $country)) . $eol;
            $content .= 'ONF_TEL = ' . $billingAddress->getTelephone() . $eol;

            // SHIPPING ADDRESS
            if (!$isB2cOrder) {
                $content .= 'ONF_NOM_L = ' . strtoupper(
                        str_replace($toReplaceList, $replacementList, $billingAddress->getCompany())
                    ) . $eol;
            }
            else {
                $content .= 'ONF_NOM_L = ' . strtoupper(
                        str_replace($toReplaceList, $replacementList, $shippingAddress->getName())
                    ) . $eol;
            }
            $streetArray = $shippingAddress->getStreet();
            $address1    = isset($streetArray[0]) ? $streetArray[0] : "";
            $address2    = isset($streetArray[1]) ? $streetArray[1] : "";
            $city        = $shippingAddress->getCity();
            $state       = $shippingAddress->getRegion();
            $country     = $this->getCountryName($shippingAddress->getCountryId());

            $content .= 'ONF_ADR1_L = ' . str_replace($toReplaceList, $replacementList, $address1) . $eol;
            $content .= 'ONF_ADR2_L = ' . str_replace($toReplaceList, $replacementList, $address2) . $eol;
            $content .= 'ONF_CP_L = ' . $shippingAddress->getPostcode() . $eol;
            $content .= 'ONF_VILLE_L = ' . strtoupper(str_replace($toReplaceList, $replacementList, $city)) . $eol;
            $content .= 'ONF_ETAT_L = ' . str_replace($toReplaceList, $replacementList, $state) . $eol;
            $content .= 'ONF_PAYS_L = ' . strtoupper(str_replace($toReplaceList, $replacementList, $country)) . $eol;
            $content .= 'ONF_TEL_L = ' . $shippingAddress->getTelephone() . $eol;

            // ORDER LINES
            /** @var \Magento\Sales\Model\Order\Item $item */
            foreach ($items as $item) {
                $options   = $item->getProductOptions();
                $jewelType = "";

                // Load product
                $sku = $item->getSku();
                if (empty($sku)) {
                    continue;
                }

                try {
                    $product = $this->_productRepository->getById($item->getProductId());
                    if ($product->getId()) {
                        $jewelType = $product->getResource()->getAttribute('jewel_type')->setStoreId(0)->getFrontend()
                                             ->getValue($product);
                    }
                } catch (NoSuchEntityException $e) {
                    continue;
                }

                $content .= '[LIGNE]' . $eol;
                $content .= 'REF FOU = ' . $sku . $eol;
                $content .= 'QTE = ' . $item->getQtyOrdered() . $eol;
                $content .= 'PA HT = ' . $item->getRowTotal() . $eol;

                if (isset($options["options"])) {
                    /*
                     * Use "getAttributeRawValue" istead of "$item->getProduct()->getData('product_custom_attribute')"
                     * to avoid loading the whole product just for this attribute.
                     * This also makes sure that we always get the admin value, which is what's needed, if the attribute
                     * scope moves from "Global" to "Store view".
                     */
                    $productCustomAttribute = $this->getProductResourceModel()
                                                   ->getAttributeRawValue(
                                                       $item->getProductId(),
                                                       'product_custom_attribute',
                                                       0
                                                   );
                    if (!empty($productCustomAttribute)) {
                        $itemName = $productCustomAttribute;
                    }
                    else {
                        $itemName = $item->getName();
                    }

                    $content .= 'LIBELLE = ' . strtoupper(str_replace($toReplaceList, $replacementList, $itemName)) . ' ';
                    foreach ($options["options"] as $option) {
                        $content .= strtoupper(str_replace($toReplaceList, $replacementList, $option['value'])) . ' ';
                    }

                    $content                 .= $eol;
                    $needExtraCarriageReturn = false;
                    foreach ($options["options"] as $option) {
                        if ($jewelType != "Bracelet" && in_array($option ['value'], $sizeList)) {
                            $content                 .= 'TAILLE = ' . trim(
                                    str_replace(["mm", "cm"], ["", ""], $option ['value'])
                                ) . ' ';
                            $needExtraCarriageReturn = true;
                        }
                    }

                    if ($needExtraCarriageReturn) {
                        $content .= $eol;
                    }
                }
            }

            $content .= '[FIN]' . $eol;
        }

        $this->_file->write($filePath, $content);

        return $filePath;
    }

    /**
     * @return \Magento\Catalog\Model\ResourceModel\Product
     */
    public function getProductResourceModel(): \Magento\Catalog\Model\ResourceModel\Product
    {
        return $this->productResourceModel;
    }
}
