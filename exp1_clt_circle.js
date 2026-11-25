// ============================================================
// exp1_clt_circle.js
// 变化定位任务 — 单圆等距布局 + 自绘虚拟光标版（2026-11-19）
// 由 exp1_clt.js 派生；原版保持不变。两项差异：
//   ① 色块中心落在「以画布短边为直径、注视点为圆心」的圆上，
//      四象限均匀、避开象限坐标轴线；所有色块到注视点等距。
//   ② 探测屏隐藏真实光标，自绘虚拟光标从注视点出发；
//      仍用鼠标左键点击认为发生了变化的那个色块。
// 基于 CLT.m (0630版) 实验规则
//
// 实验设计：
//   3×4 设计：set size (3, 4, 6) × 记忆阵列呈现时长 (50, 200, 500, 1000 ms)
//   12 种条件各 30 trial，共 360 trial，整体随机后均分到 9 个 block（每 block 40 trial）
//   练习 15 trial（3 种 set size × 5），每个试次随机分配一档呈现时长
//
// 试次结构：
//   1. ITI              (500 ms)           显示注视点
//   2. 学习阵列         (50/200/500/1000 ms)      显示注视点，时长由当前试次条件决定
//   3. 保留间隔         (1000 ms)          显示注视点（注视点全程固定显示，不消失）
//   4. 测试阵列         (最多 4000 ms)      超时未反应记为 null
//      自绘虚拟光标从注视点出发，被试移动鼠标后左键点击变化的色块
//
// Block 规则：
//   9 个 block，每 block 40 trial；block 间强制休息 15 秒（含倒计时），
//   倒计时结束后显示「您可以继续休息或按【SPACE】开始下一组实验」，按 SPACE 进入下一组
//   不设「低正确率自动结束」：无论正确率高低，实验都完整跑完 9 个 block
//   低正确率只做标记（详见下方「低正确率标记」段）
//
// 调试模式（被试信息四个输入框全部填 1，隐藏触发）：
//   正式实验只运行 1 个 block、5 个 trial；练习仍为 15 trial，跳过 block 间休息
//   被试编号与随机种子固定为 1111，导出文件名固定为 exp1_clt_1111.xlsx
//
// 低正确率标记（导出文件名前加 F，不改变实验流程）：
//   判定条件：set size = 3 或 4 时，「呈现时长 ≥ 200 ms」试次的正确率低于 50%
//   （set size 3 和 4 分别计算，任一低于 50% 即触发；set size = 6 不参与判定）
//   触发的条件行会在 K 值表中以「LowAccRed = 1」标记（标红功能受 XLSX 版本所限，见下）
//
// 导出（Excel，3 个 sheet，调试模式与正式实验规则完全一致）：
//   文件名：实验名称_被试编号（exp1_clt_编号）；
//           存在低正确率条件时在最前面加 F（如 Fexp1_clt_0112345678.xlsx）；
//           调试模式不加 F（实验名称仍为 exp1_clt）
//   Sheet1「Trial Data」：全部 trial，按实验实际发生顺序排列
//                        （含计划呈现时长 studyDuration 与实测值 studyDurationActual）
//                        超时试次（4000 ms 内未点击）以 TrialStatus = 'timeout'、
//                        CountedInK = 0 明确标记；正常反应为 'valid' / 1
//   Sheet2「K Values」：按「set size × 呈现时长」拆分的 12 个条件各一行，
//                      给出正确率、正确试次平均反应时与容量 K；
//                      末尾再附每个 set size 的汇总行（StudyDuration 记为「汇总」）；
//                      超时试次不纳入 K 计算，数量见 TotalTrials / TimeoutTrials / ValidTrials；
//                      低正确率条件行以 LowAccRed = 1 标记（标红受 XLSX 版本限制）
//   Sheet3「Grouped Trials」：列与 Sheet1 相同，但按「呈现时长 → set size」分组重排，
//                            同一条件的试次在表中连续排列，便于逐条件核对原始数据
//   K = (Acc * N^2 - N) / (N - 1)
//
// 标红能力说明：
//   当前 vendor 内嵌的 SheetJS 0.18.5 不支持单元格样式写入，
//   K 值表用「LowAccRed」列记录判定结果（1 = 应标红），将来升级 XLSX 或
//   改用 ExcelJS 后 markKRowsRed() 会自动生效，调用方无需改动。
//
// 颜色规则（每个试次由 buildProbePlan 一次算好并冻结，drawPlanSquares 只按计划绘制）：
//   记忆阵列：从 9 色调色板中不重复地随机取 setSize 个颜色
//   探测阵列：复制记忆阵列颜色，只替换「变化方块」一个 → 有且只有一个色块变色
//   新颜色  ：只能从「记忆阵列中没出现过的颜色」里挑（严格排除，无兜底放宽）
//   计划生成后立即自检 verifyProbePlan()，违反则 console.error 报警（probePlanViolations）
//   探测阵列直接使用记忆阵列那一份 plan，不会因重入/并发读到别的试次的数据
//
// 鼠标反应：只用鼠标左键点击探测阵列中的色块（circle 版）
//   只有左键点击落在色块范围内才算作反应。
//   命中坐标来自自绘虚拟光标（探测屏期间真实光标隐藏，光标从注视点出发），
//   不再使用 e.clientX/clientY。
//   实验画布仍禁用右键菜单，仅用于防止菜单遮挡刺激，与作答无关。
//
// 变化颜色规则：
//   被选方块变为一个未在阵列中出现过的新颜色。
// ============================================================

// 实验名称：仅用于导出 Excel 的文件名前缀（exp1_clt_被试编号.xlsx）
// 改名只需改这一处，文件名与结束页的手动下载链接会同步生效
const EXPERIMENT_NAME = 'exp1_clt_circle';
const EXPERIMENT_VERSION = 'clt_circle';
const SESSION_ID = (globalThis.crypto && typeof globalThis.crypto.randomUUID === 'function')
    ? globalThis.crypto.randomUUID()
    : 'session-' + Date.now() + '-' + Math.random().toString(16).slice(2);

// 运行版本标记（与 exp1_clt.html 中 exp1_clt.js?v= 保持一致）：
// 用于确认浏览器实际加载的是最新代码，而非缓存里的旧版本
const EXP1_CLT_BUILD = '20261119circle7';
console.log('[exp1_clt] build ' + EXP1_CLT_BUILD + ' loaded');

// 页面上的 build 角标已移除（避免被试在设置页看到版本号），
// 版本核对改看下面这行控制台输出，以及与 html 的 PAGE_BUILD 不一致时的告警
console.log('[exp1_clt] build ' + EXP1_CLT_BUILD + ' loaded');

// 版本自检：页面（exp1_clt.html）与脚本（exp1_clt.js）版本号必须一致。
// 不一致说明浏览器缓存了旧的 exp1_clt.js（实验会跑出旧规则、旧导出文件名），
// 在设置页直接给出告警，避免又用旧代码跑一整轮实验。
(function checkBuildMatch() {
    const pageBuild = window.EXP1_CLT_PAGE_BUILD;
    if (pageBuild && pageBuild !== EXP1_CLT_BUILD) {
        console.warn('[exp1_clt] 版本不一致：页面 ' + pageBuild + ' / 脚本 ' + EXP1_CLT_BUILD +
            ' —— 浏览器可能缓存了旧 exp1_clt.js，请 Ctrl+F5 强制刷新');
        const el = document.getElementById('build-mismatch');
        if (el) el.classList.remove('hidden');
    }
})();

const storage = window.CLTStorage;
let p = {};
let stimData = [];
let practiceData = [];
let currentBlock = 1;
let currentTrial = 0;
let blockTrials = [];
let experimentPhase = 'setup';

// --- 断点续做：PRNG 基准与续做状态（2026-10-10，与 exp1_fcdt 结构一致）---
// buildFormalTrials 每次都先把 PRNG 重置回正式试次基准起点，
// 这样无论是否跳过练习，360 试次的顺序都完全一致。
let formalPrngBaseSeed = 0;
let isResumedSession = false;
let resumedFromBlock = 0;   // 0 = 非续做；N = 从第 N 组（block）续做
let resumeCount = 0;        // 本被试累计续做次数（第 1 次续做 = 1）
// 续做起点：若检测到上次未完成的进度，从第 N 组（block）继续；首次（或重新开始）为 1。
let resumeStartBlock = 1;
let currentNItems = 4; // 当前试次的 set size，用于鼠标命中范围判断

// --- 自绘虚拟光标状态（circle 版）---
let virtualCursor = { x: 0, y: 0 };
let cursorActive = false;
let cursorSkipNextMove = false;
let cursorStartX = 0, cursorStartY = 0;
let cursorTravelPx = 0;
let lastClientX = null, lastClientY = null;
let pointerLocked = false;
let probeSnapshot = null;            // { locs, colorIdx }：探测屏重绘用
let clickedItemIndexCircle = -1;     // 本试次点中的色块序号（0 起）
let lowAccBlockCount = 0;        // 正确率低于 50% 的 block 累计数（只统计，不终止实验）
// 低正确率标记：set size = 4 的平均正确率低于 50% 时为 true，导出文件名最前面加 F。
// 注意：实验本身不再因低正确率提前结束，永远跑完 9 个 block。
let ss4LowAccuracyFlag = false;  // 实际语义：低正确率标记（SS3/SS4 在 ≥200ms 正确率 < 50%）
let lowAccDetail = { triggered: false, details: [] };  // 各 set size 的判定明细，供 K 值表标红使用

// --- 实验参数（对应 CLT.m getPreferences） ---
const prefs = {
    setSizes: [3, 4, 6],
    studyDurations: [50, 200, 500, 1000],
    numBlocks: 9,
    numTrials: 40,
    nEachCondition: 30,      // 每个 set size × 呈现时长组合 30 trial
    nPractice: 15,          // 练习 15 trial，3 种 set size 各 5 trial
    ITI: 500,               // ms
    retentionInterval: 1000,// ms
    stimSize: 51,
    minDist: 51 * 2.5,
    fixationSize: 8,
    breakLength: 0.25,      // 分钟 = 15 秒（block 间强制休息，与 exp1_fcdt 保持一致）
    labelFontSize: 18
};

// --- circle 版专属参数（与 exp1_fcdt_circle 保持一致）---
const CIRCLE_EDGE_MARGIN = 8;    // px：名义参考圆内缩留白
// 经典变化觉察/圆形阵列研究常用约 7.3° 的垂直刺激区域、约 1.95° 的圆周半径；
// 这里用 1.95 / 7.3 ≈ 0.267 将半径映射到画布短边，避免刺激离注视点过远。
// 统一刺激圆外半径：参考坐标系 768 px 下为 200 px，按画布短边同比例缩放。
const STIMULUS_CIRCLE_RADIUS_RATIO = 200 / 768;
const CIRCLE_RADIUS_RATIO = STIMULUS_CIRCLE_RADIUS_RATIO;
const AXIS_MARGIN_DEG = 15;      // 度：避开 0/90/180/270° 象限坐标轴线
const TRY_POINTER_LOCK = true;   // 探测前锁定指针，保证虚拟光标可覆盖整个画布

// 调试模式只缩短实际运行的正式实验；prefs 始终保留完整科研设计参数，便于自检。
let runtimeNumBlocks = prefs.numBlocks;
let runtimeNumTrials = prefs.numTrials;
let runtimeBreakLength = prefs.breakLength;

// --- 兼容旧版数据结构的内部标签映射 ---
// 鼠标版不再绘制数字标签，但保留映射字段，便于旧表结构继续接收 response/correctResponse。
const LABEL_CHARS = ['1', '2', '3', '8', '9', '0'];
let trialLabels = [];

const colors_9 = [
    [255, 0, 0], [0, 255, 0], [0, 0, 255], [255, 255, 0],
    [255, 0, 255], [0, 255, 255], [255, 255, 255], [0, 0, 0], [255, 128, 0]
];

// --- PRNG ---
let seed = 1;
function seedRandom(s) { seed = s; }
function random() {
    let x = Math.sin(seed++) * 10000;
    return x - Math.floor(x);
}
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}
function randInt(min, max) {
    return Math.floor(random() * (max - min + 1)) + min;
}

// --- 高精度计时器 ---
function wait(ms) {
    return new Promise(resolve => {
        const start = performance.now();
        function check(time) {
            if (time - start >= ms) { resolve(); }
            else { requestAnimationFrame(check); }
        }
        requestAnimationFrame(check);
    });
}

// --- DOM 元素 ---
// 问卷阶段已删除，因此 screens 中不再包含 questionnaire 键
const screens = {
    setup: document.getElementById('setup-phase'),
    instructions: document.getElementById('instructions-phase'),
    experiment: document.getElementById('experiment-phase'),
    break: document.getElementById('break-phase'),
    upload: document.getElementById('upload-phase'),
    finished: document.getElementById('finished-phase')
};
const canvas = document.getElementById('expCanvas');
const ctx = canvas.getContext('2d', { alpha: false });

// --- 身份证号掩蔽 & SubjectID 自动生成 ---
let realIdCard = ''; // 存储真实身份证号（掩蔽显示时仍保留原值）

const idCardInput = document.getElementById('subIdCard');
const phoneInput = document.getElementById('subPhone');
const subjectIDInput = document.getElementById('subjectID');

// 身份证号：输入时更新真实值
idCardInput.addEventListener('input', () => {
    realIdCard = idCardInput.value;
    updateSubjectID();
});
// 聚焦时恢复真实值便于编辑
idCardInput.addEventListener('focus', () => {
    idCardInput.value = realIdCard;
});
// 失焦时掩蔽显示（前4位 + * + 后6位）
idCardInput.addEventListener('blur', () => {
    realIdCard = idCardInput.value;
    if (realIdCard.length > 10) {
        idCardInput.value = realIdCard.substring(0, 4)
            + '*'.repeat(realIdCard.length - 10)
            + realIdCard.substring(realIdCard.length - 6);
    }
});

// 手机号变化时更新 SubjectID
phoneInput.addEventListener('input', updateSubjectID);
// 调试模式复选框监听：勾选时自动填充默认调试信息
const debugModeCheckbox = document.getElementById('debugMode');
if (debugModeCheckbox) {
    debugModeCheckbox.addEventListener('change', (e) => {
        if (e.target.checked) {
            const nameEl = document.getElementById('subName');
            const genderEl = document.getElementById('subGender');
            const ageEl = document.getElementById('subAge');
            if (nameEl) nameEl.value = '1';
            if (genderEl) genderEl.value = '女';
            if (ageEl) ageEl.value = '1';
            realIdCard = '111111111111111111';
            if (idCardInput) idCardInput.value = '1111********111111';
            if (phoneInput) phoneInput.value = '11111111111';
            if (subjectIDInput) subjectIDInput.value = '1111';
        }
    });
}

function updateSubjectID() {
    const id = buildSubjectID();
    subjectIDInput.value = id;
}

function buildSubjectID() {
    const idVal = realIdCard || idCardInput.value || '';
    const phoneVal = phoneInput.value || '';
    if (idVal.length >= 6 && phoneVal.length >= 4) {
        return idVal.slice(-6) + phoneVal.slice(-4);
    }
    return '';
}

// 实验日期（YYYYMMDD），随数据一起保存；不再用于导出文件名
function formatDateYYYYMMDD(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${day}`;
}

// 掩蔽身份证号（用于数据导出，保护隐私）
// 空值安全：导出时可能为未填写/未初始化，此时返回空串而不是抛错
function maskIdCard(idCard) {
    if (!idCard) return '';
    if (idCard.length > 10) {
        return idCard.substring(0, 4)
            + '*'.repeat(idCard.length - 10)
            + idCard.substring(idCard.length - 6);
    }
    return idCard;
}

// 全角数字 → 半角数字（中文输入法打出的「１」会被折算成「1」）
// 组别校验与调试模式触发判定共用，避免「明明填了却判不中」。
function normalizeDigits(v) {
    return String(v == null ? '' : v)
        .trim()
        .replace(/[\uFF10-\uFF19]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30));
}

// 组别（两个范式的先后顺序，被试间变量）：
//   1 = 先做变化觉察任务（FCDT）再做变化定位任务（CLT）
//   2 = 先做变化定位任务（CLT）再做变化觉察任务（FCDT）
// 两个实验程序共用同一套编号，导出列名为 Order，合并两个范式的数据后
// 即可检验「Order × ExperimentType」的交互（即顺序效应）。
const GROUP_LABELS = {
    1: '1 = 先做变化觉察任务（FCDT），再做变化定位任务（CLT）',
    2: '2 = 先做变化定位任务（CLT），再做变化觉察任务（FCDT）'
};

function readSubGroup() {
    const input = document.getElementById('subGroup');
    return normalizeDigits(input ? input.value : '');
}

// --- 设置表单 ---
document.getElementById('start-btn').addEventListener('click', async () => {
    const form = document.getElementById('setup-form');
    if (!form.reportValidity()) return;

    // 组别：两个范式的先后顺序（被试间变量），只接受 1 或 2。
    // HTML 侧已用 pattern="[12]" + required 做第一层拦截，这里再显式校验一次：
    // 一是兜住全角数字与粘贴带空格的情况，二是把合法取值告诉被试（页面上不写）。
    const subGroupValue = readSubGroup();
    if (!/^[12]$/.test(subGroupValue)) {
        alert('组别填写有误，请按主试安排填写：\n' + GROUP_LABELS[1] + '\n' + GROUP_LABELS[2]);
        const subGroupInput = document.getElementById('subGroup');
        if (subGroupInput && subGroupInput.focus) subGroupInput.focus();
        return;
    }
    p.subGroup = Number(subGroupValue);

    p.subName = document.getElementById('subName').value;
    p.subGender = document.getElementById('subGender').value;
    p.subAge = document.getElementById('subAge').value;
    p.subIdCard = realIdCard || idCardInput.value; // 使用真实身份证号
    p.subPhone = document.getElementById('subPhone').value;
    // 被试编号：点击开始时按规则由当前输入框的值重新生成一次
    // （不依赖 input 事件，浏览器自动填充 / 程序赋值也能正确生成）
    if (buildSubjectID()) subjectIDInput.value = buildSubjectID();
    // 严格按「身份证后6位 + 手机后4位」的原始数字字符串保存，不做 parseInt
    // （否则前导 0 会被吞掉，如 0701231234 变成 701231234）
    const generatedID = String(subjectIDInput.value || '').trim();
    p.subjectID = generatedID || String(Date.now());
    if (!generatedID) {
        // 身份证号不足 6 位或手机号不足 4 位时无法按规则生成编号，
        // 只能退回时间戳。明确告警，避免事后把「一串数字」误当成随机种子。
        console.warn('[exp1_clt] 被试编号无法按规则生成（身份证号需 ≥6 位、手机号需 ≥4 位），' +
            '已退回时间戳：' + p.subjectID);
    }
    // 随机种子：用字符串编号派生（Number 转换仅用于 PRNG 内部，不影响编号本身）
    const seedNum = Number(p.subjectID);
    p.rndSeed = Number.isFinite(seedNum) ? seedNum : Date.now();
    p.saveLocal = true; // 保存到本地始终开启（复选框已隐藏，网页端数据库 + 本地均保存）
    p.expDate = formatDateYYYYMMDD(new Date()); // 实验日期（随数据保存，不参与文件名）

    // 调试模式触发（复选框已隐藏）：姓名/年龄/身份证号/手机号 四个输入框全部填 1
    // （性别为下拉选择，不参与判定；subjectID 自动生成不受影响）
    // 判定做了两处容错，避免「明明填了 1 却没进调试模式、导出文件名变成一串时间戳」：
    //   1) 中文输入法打出的全角「１」自动折算为半角；
    //   2) 只要四个框的内容「全部由数字 1 组成」即算命中 —— 身份证号框受 maxlength=18
    //      限制，连打 18 个 1 时也是命中；且用的是未掩蔽的 realIdCard，
    //      不会因为失焦掩蔽成「1111********111111」而漏判。
    // 全角折算复用 normalizeDigits()（与组别校验同一个 helper）
    const isAllOnes = (v) => /^1+$/.test(normalizeDigits(v));
    p.debugMode = isAllOnes(p.subName);
    if (p.debugMode) {
        console.log('[exp1_clt] 调试模式已触发（姓名/年龄/身份证号/手机号 全为 1）');
    } else {
        console.log(`[exp1_clt] 正式模式（姓名="${p.subName}" 年龄="${p.subAge}" ` +
            `身份证号长度=${String(p.subIdCard || '').length} 手机号="${p.subPhone}"）`);
    }

    // 调试模式：练习仍使用完整的 15 个试次；正式实验缩短为 1 个 block × 5 trial，
    // 并跳过 block 间休息。prefs 保留完整设计参数，避免调试设置干扰正式试次生成自检；
    // 被试编号与随机种子一并固定为 1111（导出文件名为 exp1_clt_1111.xlsx），便于识别与复现。
    runtimeNumBlocks = prefs.numBlocks;
    runtimeNumTrials = prefs.numTrials;
    runtimeBreakLength = prefs.breakLength;
    if (p.debugMode) {
        runtimeNumBlocks = 1;
        runtimeNumTrials = 5;
        runtimeBreakLength = 0;
        p.subjectID = '1111';
        p.rndSeed = 1111;
        subjectIDInput.value = '1111'; // 设置页同步显示，与导出文件名中的被试编号一致
    }

    seedRandom(p.rndSeed);

    // 正式试次序列使用专属 PRNG 种子，与练习阶段完全解耦（与 exp1_fcdt 一致）：
    // 无论练习是否执行 / 跳过，buildFormalTrials 输出的试次顺序恒定，
    // 续做时跳过练习直接从第 N 组继续，被试看到的 360 试次序列与首次完全一致。
    formalPrngBaseSeed = p.rndSeed + 7919;

    // 该范式的虚拟光标依赖 Pointer Lock 覆盖整屏；锁定失败时不降级，
    // 避免真实鼠标碰到物理屏幕边缘后无法继续移动。
    if (!canvas.requestPointerLock) {
        showSetupError('当前浏览器不支持鼠标指针锁定，请使用最新版 Chrome 或 Edge。');
        return;
    }
    let lockPromise;
    try {
        lockPromise = canvas.requestPointerLock();
    } catch (err) {
        showSetupError('无法锁定鼠标指针，请允许浏览器锁定鼠标后重试。');
        return;
    }
    let fullscreenPromise = null;
    try {
        if (document.documentElement.requestFullscreen) {
            fullscreenPromise = document.documentElement.requestFullscreen();
        }
    } catch (err) { console.warn('Fullscreen denied.', err); }
    if (lockPromise && typeof lockPromise.then === 'function') {
        try { await lockPromise; } catch (err) {
            showSetupError('浏览器未允许锁定鼠标指针，实验未开始。请重新点击开始并允许锁定。');
            return;
        }
    }
    // Pointer Lock 状态通过 pointerlockchange 异步更新，不能立即读取。
    if (document.pointerLockElement !== canvas) {
        const locked = await new Promise(resolve => {
            let settled = false;
            const finish = value => {
                if (settled) return;
                settled = true;
                document.removeEventListener('pointerlockchange', onChange);
                clearTimeout(timer);
                resolve(value);
            };
            const onChange = () => finish(document.pointerLockElement === canvas);
            const timer = setTimeout(() => finish(false), 1500);
            document.addEventListener('pointerlockchange', onChange);
        });
        if (!locked) {
            showSetupError('鼠标指针锁定未生效，实验未开始。请重新点击开始并允许锁定。');
            return;
        }
    }
    if (fullscreenPromise && typeof fullscreenPromise.then === 'function') {
        try { await fullscreenPromise; } catch (err) { console.warn('Fullscreen denied.', err); }
    }

    try {
        if (!storage) throw new Error('Storage adapter is unavailable');
        // experiment 决定 Batch Session 的进度键名（exp1_clt:<workerId>），
        // 必须传，否则续做会退化成所有实验共用一个键
        await storage.initializeStorage({
            participant: { ...p }, build: EXP1_CLT_BUILD, experiment: EXPERIMENT_NAME
        });
    } catch (err) {
        console.error('[CLT storage] initialization failed:', err);
        if (storage && storage.backendName === 'jatos') {
            showSetupError('JATOS 初始化失败，实验尚未开始。请联系实验管理员。');
            return;
        }
    }

    // 检查是否有可续做的进度（仅 JATOS Batch Session 后端提供；其它后端返回 null，
    // 等价于"从头开始"，不影响正常流程）。续做从上次完成后的下一组继续，
    // 试次序列因 PRNG 基准固定而与首次完全一致。
    // 归属校验（这条进度是否属于当前被试）已在存储层完成：不属于、或无法校验时返回 null。
    let resumable = null;
    try {
        resumable = await storage.loadProgress();
    } catch (err) {
        console.warn('[' + EXPERIMENT_NAME + '] 读取续做进度失败，将从头开始：', err);
    }
    if (resumable && resumable.completedBlocks > 0 && resumable.completedBlocks < runtimeNumBlocks) {
        resumeStartBlock = resumable.completedBlocks + 1;
        isResumedSession = true;
        resumedFromBlock = resumeStartBlock;
        // 本次续做计入累计次数：进度里记的是"此前续做过几次"，本次 +1 后再随进度
        // 写回；导出时该被试的每一行都带同一数值（一次做完的被试为 0）
        resumeCount = (resumable.resumeCount || 0) + 1;
        console.log('[' + EXPERIMENT_NAME + '] 检测到未完成的进度（已完成 ' +
            resumable.completedBlocks + ' 组），将从第 ' + resumeStartBlock + ' 组续做。');
    }

    // 检测到可续做进度时，显示"继续 / 重新开始"选择条；否则隐藏。
    switchScreen('instructions');
    experimentPhase = 'instructions';

    const resumeNotice = document.getElementById('resume-notice');
    if (resumeNotice) {
        if (resumeStartBlock > 1) resumeNotice.classList.remove('hidden');
        else resumeNotice.classList.add('hidden');
    }
    const btnResume = document.getElementById('btn-resume');
    if (btnResume) btnResume.addEventListener('click', () => {
        resumeNotice.classList.add('hidden');
        // 保持 resumeStartBlock 不变，指导语读完后 SPACE 键按续做起点分流
    });
    const btnRestart = document.getElementById('btn-restart');
    if (btnRestart) btnRestart.addEventListener('click', () => {
        resumeStartBlock = 1;
        isResumedSession = false;
        resumedFromBlock = 0;
        resumeCount = 0;
        if (storage) storage.clearProgress();
        resumeNotice.classList.add('hidden');
        console.log('[' + EXPERIMENT_NAME + '] 被试选择从头重新开始，已清除线上进度。');
    });
});

function showSetupError(message) {
    const el = document.getElementById('build-mismatch');
    if (!el) return;
    el.innerText = message;
    el.classList.remove('hidden');
}

function switchScreen(screenName) {
    for (let key in screens) screens[key].classList.add('hidden');
    screens[screenName].classList.remove('hidden');
}

// --- 全局键盘监听：仅用于阶段切换，不再用于试次作答 ---
window.addEventListener('keydown', (e) => {
    if (e.repeat) return;

    if (experimentPhase === 'instructions' && (e.code === 'Space' || e.key === ' ')) {
        e.preventDefault();
        // 续做（resumeStartBlock > 1）：被试已在首次做过练习，跳过直接进正式；
        // 否则正常走练习 → 正式流程。
        if (resumeStartBlock > 1) startFormal();
        else startPractice();
    } else if (experimentPhase === 'practice_end' && (e.code === 'Space' || e.key === ' ')) {
        e.preventDefault();
        startFormal();
    } else if (experimentPhase === 'break_countdown' || experimentPhase === 'break_wait') {
        // 休息阶段：强制倒计时期间按 SPACE 无效；
        // 倒计时结束后按 SPACE 进入下一组（无「继续休息」按钮，也无 ESC 退出）
        if (experimentPhase === 'break_wait' && (e.code === 'Space' || e.key === ' ')) {
            e.preventDefault();
            nextBlock();
        }
    }
});

// ============================================================
//  防误触层（guard rails）
//  目的：避免被试或浏览器在实验进行中触发以下操作导致实验中断 / 数据丢失：
//   1) 浏览器刷新、关闭、离开（beforeunload）
//   2) 浏览器后退 / 前进（Backspace、Alt+← / Alt+→、popstate）
//   3) 快捷键被劫持（F5 / Ctrl+R / Ctrl+W / Ctrl+T / F12 / DevTools）
//   4) Esc 退出全屏（屏蔽 Esc）
//   5) 鼠标长按 → 系统右键菜单 / 拖拽（mousedown 持续触发、dragstart、selectstart）
//   6) 触屏 / 滚轮 / 缩放（touchstart、wheel、gesturestart）
//   7) 页面失焦 / 切到后台（visibilitychange 给出告警但不清空数据）
//   8) 页面被 resize（仅在 #experiment-phase 锁住尺寸，避免破坏注视点位置）
//  实验进行中：仅按每个阶段规定的"有效键"（SPACE / 鼠标左键）才生效；
//  其他任何键 / 鼠标 / 触屏 / 系统事件一律被静默 preventDefault，实验流程不受影响。
// ============================================================

// 每个阶段被允许的键盘按键集合（其他键一律屏蔽）
const ALLOWED_KEYS = new Set([' ', 'Space', 'Spacebar']);

// 实验中需要拦截的系统级组合键（不拦截时浏览器会触发默认行为）
const BLOCKED_SHORTCUTS = new Set([
    'F5', 'F12',                                                  // 刷新 / 开发者工具
    'F1', 'F2', 'F3', 'F4', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11' // 其他功能键
]);
function isBlockedShortcut(e) {
    if (BLOCKED_SHORTCUTS.has(e.key)) return true;
    if (e.ctrlKey || e.metaKey) {
        // 拦截 Ctrl/Cmd + R / W / T / N / P / F / S / Shift+Delete 等
        if (/^[rwnpfst]$/i.test(e.key)) return true;
        if (e.key === 'F4' && e.altKey) return true;   // Alt+F4 关窗
    }
    if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.key === 'Backspace')) {
        return true;                                       // Alt+←/→ 浏览器后退 / 前进
    }
    return false;
}

// (1) 全局 keydown 第二层：仅放行 ALLOWED_KEYS，其余静默屏蔽
window.addEventListener('keydown', (e) => {
    if (e.repeat) return;                                 // 防止长按重复触发

    // 关键例外：焦点落在输入框 / 下拉框里时（setup 阶段填姓名 / 性别 / 年龄 /
    // 身份证 / 手机号），防误触层必须放行普通按键——在 capture 阶段对 keydown
    // 调用 preventDefault() 会阻止字符被插入输入框，被试将一个字都打不进去。
    const inFormField = (e.target instanceof HTMLInputElement)
        || (e.target instanceof HTMLTextAreaElement)
        || (e.target instanceof HTMLSelectElement);

    if (isBlockedShortcut(e)) {
        e.preventDefault();
        e.stopPropagation();
        return;
    }
    if (inFormField) return;                               // 表单输入：放行所有普通按键
    if (e.key === 'Escape') {                              // 屏蔽 Esc 退出全屏
        e.preventDefault();
        return;
    }
    if (e.key === 'Backspace' && !(e.target instanceof HTMLInputElement)) {
        // Backspace 在非输入框下被浏览器解读为"后退" → 屏蔽
        e.preventDefault();
        return;
    }
    // 其他所有非放行键：静默屏蔽
    if (!ALLOWED_KEYS.has(e.key) && !ALLOWED_KEYS.has(e.code) && e.key !== 'Tab') {
        e.preventDefault();
    }
}, true);                                                  // capture 阶段优先

// (2) 屏蔽所有元素的右键菜单（不仅 canvas）
//     实验阶段 canvas 上的右键必须屏蔽（菜单会遮挡刺激、干扰作答）；
//     但 setup 阶段表单里的右键菜单保留（便于粘贴身份证号 / 手机号），
//     文本输入框的右键菜单里没有任何导航项，不会导致实验中断。
window.addEventListener('contextmenu', (e) => {
    const t = e.target;
    if (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement) return;
    e.preventDefault();
});

// (3) 屏蔽拖拽 / 选中 / 复制（避免误选文字拖出浏览器）
window.addEventListener('dragstart', (e) => e.preventDefault());
window.addEventListener('selectstart', (e) => {
    // 允许在 setup 阶段（输入框）选中文字；其他阶段屏蔽
    const phase = experimentPhase;
    if (phase !== 'setup' && phase !== 'instructions') {
        if (!(e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement)) {
            e.preventDefault();
        }
    }
});
window.addEventListener('copy', (e) => {
    if (!(e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement)) {
        e.preventDefault();
    }
});

// (4) 屏蔽滚轮 / 触屏手势（避免页面缩放 / 滚动 / 系统菜单）
window.addEventListener('wheel', (e) => {
    if (e.ctrlKey) e.preventDefault();                    // Ctrl+滚轮 = 缩放
}, { passive: false });
window.addEventListener('touchstart', (e) => {
    if (e.touches.length > 1) e.preventDefault();         // 多指 = 缩放手势
}, { passive: false });
window.addEventListener('gesturestart', (e) => e.preventDefault());

// (5) 屏蔽全屏切换 / 重新加载（beforeunload）：实验进行中给出离开提示
window.addEventListener('beforeunload', (e) => {
    const phase = experimentPhase;
    if (phase === 'setup' || phase === 'finished') return;  // 未开始 / 已结束：允许直接离开
    e.preventDefault();
    e.returnValue = '实验正在进行中，关闭或刷新页面将导致数据丢失。确定要离开吗？';
    return e.returnValue;
});

// (6) 页面失焦 / 切到后台：给出告警，但**不**清空数据、不终止试次
let visibilityWarned = false;
document.addEventListener('visibilitychange', () => {
    if (document.hidden && !visibilityWarned) {
        visibilityWarned = true;
        console.warn('[' + EXPERIMENT_NAME + '] 页面失焦：被试切到其他窗口 / 标签页。');
    } else if (!document.hidden) {
        visibilityWarned = false;
    }
});
window.addEventListener('blur', () => {
    if (experimentPhase === 'trial_response' || experimentPhase === 'practice_response') {
        console.warn('[' + EXPERIMENT_NAME + '] 实验进行中窗口失焦，请保持页面在前台。');
    }
});

// (7) popstate（浏览器后退 / 前进）：压入 dummy 历史项并屏蔽后退
history.pushState({ phase: 'experiment' }, '');
window.addEventListener('popstate', () => {
    const phase = experimentPhase;
    if (phase !== 'setup' && phase !== 'finished') {
        // 实验进行中：阻止后退，立即推回一项
        history.pushState({ phase: 'experiment' }, '');
        console.warn('[' + EXPERIMENT_NAME + '] 浏览器后退被屏蔽（实验进行中）。');
    }
});

// (8) 屏蔽浏览器对鼠标的默认处理（防止长按出系统菜单、拖拽光标、指针变化）。
//     mousedown 一次按下只触发一次，浏览器在按住期间不会自动 repeat mousedown；
//     右键菜单已由 (2) contextmenu preventDefault 屏蔽；
//     这里额外阻止 mousedown 的默认光标变化 / 文本选中，并屏蔽 canvas dragstart。
canvas.addEventListener('mousedown', (e) => {
    // 在所有阶段都阻止默认行为（除 setup 阶段输入框，但 canvas 不在输入框里）
    e.preventDefault();
}, true);
canvas.addEventListener('dragstart', (e) => e.preventDefault());
canvas.addEventListener('drag', (e) => e.preventDefault());

// 与作答无关，仅为防止浏览器右键菜单遮挡刺激。
// （已由上面的 (2) 全局 contextmenu 屏蔽；这里保留 canvas 上的显式监听便于阅读）

// 判断 (x, y) 落在哪个色块上；未命中返回 -1。
// circle 版：坐标来自自绘虚拟光标，不再从鼠标事件里换算 clientX/Y。
function clickedItemIndex(x, y) {
    if (!trialPlan || !Array.isArray(trialPlan.locs)) return -1;
    const half = prefs.stimSize / 2;
    for (let i = 0; i < currentNItems; i++) {
        const loc = trialPlan.locs[i];
        if (Math.abs(x - loc.x) <= half && Math.abs(y - loc.y) <= half) return i;
    }
    return -1;
}

canvas.addEventListener('mousedown', (e) => {
    const phase = experimentPhase;
    if (phase !== 'practice_response' && phase !== 'trial_response') return;
    if (e.button !== 0) return;
    e.preventDefault();

    const itemIdx = clickedItemIndex(virtualCursor.x, virtualCursor.y);
    if (itemIdx < 0) return;
    clickedItemIndexCircle = itemIdx;

    // response/correctResponse 沿用旧版标签槽编号，确保既有数据结构兼容；
    // 实际选择的是 itemIdx 对应的色块，导出时再转换为 1 起始的色块序号。
    const mappedSlot = labelMap.indexOf(itemIdx);
    const responseNum = mappedSlot >= 0 ? mappedSlot + 1 : itemIdx + 1;

    if (phase === 'practice_response') recordPracticeResponse(responseNum);
    else recordResponse(responseNum);
});

// ============================================================
//  练习阶段（15 trial，3 种 set size × 5，呈现时长逐试次随机，全部 change）
// ============================================================
let practiceTrials = [];
let practiceTrialIdx = 0;
let practiceAcc = [];

function startPractice() {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;

    const nPerSzPractice = prefs.nPractice / prefs.setSizes.length;
    practiceTrials = [];
    for (let ss of prefs.setSizes) {
        for (let i = 0; i < nPerSzPractice; i++) {
            const durationIdx = randInt(0, prefs.studyDurations.length - 1);
            practiceTrials.push({
                setSize: ss,
                studyDuration: prefs.studyDurations[durationIdx],
                change: 1
            }); // CLT 全部 change
        }
    }
    shuffleArray(practiceTrials);
    practiceTrialIdx = 0;
    practiceAcc = [];
    practiceData = [];

    switchScreen('experiment');
    runPracticeTrial();
}

let activeColors = [];
let activeLocs = [];
let trialData = {};
let practiceTrialData = {};
let responseStartTime = 0;
let responseTimeoutTimer = null;
const RESPONSE_LIMIT_MS = 4000;

// CLT 专用：变化方块索引、新颜色、兼容旧数据的内部标签映射
// 这些全局量是「当前试次」的镜像，记录数据与鼠标判断用；
// 绘制一律走 trialPlan（见 buildProbePlan），不直接读这些全局量。
let changeItemIdx = 0;
let changeColorIdx = 0;
let labelMap = [];
let correctLabel = 0;

// --- 当前试次的「探测计划」---
// 记忆阵列颜色、探测阵列颜色、位置、标签全部由它一次算好并冻结。
// 这样探测阵列必定是「它自己那次记忆阵列」的唯一一个方块变色，
// 不可能出现两个（或更多）方块同时变色。
let trialPlan = null;
let probePlanViolations = 0;   // 自检失败计数，正常情况下恒为 0

// 自检：探测阵列相对记忆阵列必须有且只有一个方块变色，
// 且变化后的颜色不能是记忆阵列中出现过的颜色
function verifyProbePlan(plan) {
    let diff = [];
    for (let i = 0; i < plan.setSize; i++) {
        if (plan.probeColors[i] !== plan.memColors[i]) diff.push(i);
    }
    const onlyOneChanged = diff.length === 1 && diff[0] === plan.changeIdx;
    const newColorNotInMemory = !plan.memColors.includes(plan.probeColors[plan.changeIdx]);
    const colorReallyChanged = plan.probeColors[plan.changeIdx] !== plan.memColors[plan.changeIdx];
    const noDupInMemory = new Set(plan.memColors).size === plan.setSize;
    const ok = onlyOneChanged && newColorNotInMemory && colorReallyChanged && noDupInMemory;
    if (!ok) {
        probePlanViolations++;
        console.error('[exp1_clt] 探测阵列自检未通过（不应发生）：', {
            diff: diff, memColors: plan.memColors, probeColors: plan.probeColors,
            changeIdx: plan.changeIdx, setSize: plan.setSize, violations: probePlanViolations
        });
    }
    return ok;
}

// 生成一次试次的完整视觉计划
//   记忆阵列：从 9 色调色板中不重复地随机取 setSize 个颜色
//   探测阵列：复制记忆阵列颜色，只把「变化方块」换成新颜色 → 恰好 1 个方块变色
//   新颜色  ：只能从「记忆阵列中没出现过的颜色」里挑，严格排除，无兜底放宽
function buildProbePlan(setSize) {
    let colorIndexArr = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    shuffleArray(colorIndexArr);
    const memColors = colorIndexArr.slice(0, setSize);

    const locs = getStimLocs(setSize, canvas.width / 2, canvas.height / 2);

    // 变化方块：随机 1 个
    const changeIdx = randInt(0, setSize - 1);

    // 新颜色：排除记忆阵列出现过的全部颜色
    let newColorPool = [];
    for (let c = 0; c < colors_9.length; c++) {
        if (!memColors.includes(c)) newColorPool.push(c);
    }
    if (newColorPool.length === 0) {
        // 调色板共 9 色、setSize 最大 6，实际不可能触发；
        // 保留防御但不放宽条件（绝不退回记忆阵列已用的颜色）
        throw new Error('buildProbePlan: 调色板颜色不足，无法选出不属于记忆阵列的新颜色');
    }
    const newColor = newColorPool[randInt(0, newColorPool.length - 1)];

    // 探测阵列颜色：仅替换变化方块那一项
    const probeColors = memColors.slice();
    probeColors[changeIdx] = newColor;

    // 内部标签映射：鼠标版不绘制标签，仅用于保持旧数据字段的编号语义。
    let lmap = [];
    for (let i = 0; i < setSize; i++) lmap.push(i);
    shuffleArray(lmap);

    // 兼容旧数据的标签字符元数据（不显示给被试）
    let charPool = [0, 1, 2, 3, 4, 5];
    shuffleArray(charPool);

    const plan = {
        setSize: setSize,
        locs: locs,
        memColors: memColors,           // 记忆阵列颜色
        probeColors: probeColors,       // 探测阵列颜色（恰好 1 项与 memColors 不同）
        changeIdx: changeIdx,           // 变化方块的索引
        origColor: memColors[changeIdx],// 变化方块原色
        newColor: newColor,             // 变化方块新色（不在 memColors 中）
        labelMap: lmap,
        trialLabels: charPool.slice(0, setSize),
        correctLabel: lmap.indexOf(changeIdx) + 1
    };

    verifyProbePlan(plan);
    return plan;
}

// 把试次计划同步到全局镜像（记录数据 / 按键判断用；绘制不依赖这些全局量）
function applyPlanToGlobals(plan) {
    activeColors = plan.memColors.slice();
    activeLocs = plan.locs;
    changeItemIdx = plan.changeIdx;
    changeColorIdx = plan.newColor;
    labelMap = plan.labelMap.slice();
    trialLabels = plan.trialLabels.slice();
    correctLabel = plan.correctLabel;
    currentNItems = plan.setSize;
}

// 按指定颜色数组绘制当前计划的全部方块（记忆阵列 / 探测阵列共用）
function drawPlanSquares(plan, colorIdxArr) {
    for (let i = 0; i < plan.setSize; i++) {
        drawSquare(plan.locs[i].x, plan.locs[i].y, colors_9[colorIdxArr[i]]);
    }
}

async function runPracticeTrial() {
    if (practiceTrialIdx >= practiceTrials.length) {
        endPractice();
        return;
    }

    experimentPhase = 'practice_running';
    let condition = practiceTrials[practiceTrialIdx];
    currentNItems = condition.setSize;
    practiceTrialData = {
        block: 0,
        trial: practiceTrialIdx + 1,
        setSize: condition.setSize,
        studyDuration: condition.studyDuration,
        isChange: 1
    };

    // ITI
    drawBackground(); drawFixation();
    await wait(prefs.ITI);

    await showStudyArrayPractice(condition);
}

async function showStudyArrayPractice(condition) {
    // 一次算好并冻结本试次的视觉计划：记忆阵列 / 探测阵列 / 位置 / 标签
    // 计划内已保证「有且只有一个方块变色」且「新颜色不在记忆阵列中出现」
    const plan = buildProbePlan(condition.setSize);
    trialPlan = plan;
    applyPlanToGlobals(plan);

    // 学习阵列（仅方块，无标签）
    drawBackground(); drawFixation();
    drawPlanSquares(plan, plan.memColors);
    const studyStartTime = performance.now();
    await wait(condition.studyDuration);
    practiceTrialData.studyDurationActual = (performance.now() - studyStartTime).toFixed(2);

    // 保留间隔（注视点全程固定显示，不消失）
    drawBackground(); drawFixation();
    await wait(prefs.retentionInterval);

    // 探测阵列：直接用同一个 plan 的 probeColors，任何情况下都只有一个方块变色
    showTestArrayPractice(plan);
}

function showTestArrayPractice(plan) {
    // 冻结本试次探测屏内容，供 mousemove 整屏重绘时复用
    probeSnapshot = { locs: plan.locs, colorIdx: plan.probeColors };
    clickedItemIndexCircle = -1;
    resetVirtualCursor();          // 自绘光标回到注视点
    redrawProbeScreen();          // 背景 + 注视点 + 色块 + 自绘光标

    experimentPhase = 'practice_response';
    responseStartTime = performance.now();
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = setTimeout(() => {
        if (experimentPhase === 'practice_response') recordPracticeResponse(null);
    }, RESPONSE_LIMIT_MS);
}

function recordPracticeResponse(responseNum) {
    // 立即锁定反应阶段：防止鼠标连击在同一试次内触发第二次响应
    if (experimentPhase !== 'practice_response') return;
    experimentPhase = 'practice_locked';
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = null;
    hideVirtualCursor();          // 关掉自绘光标、恢复真实光标

    const timedOut = responseNum === null;
    let rt = timedOut ? null : (performance.now() - responseStartTime) / 1000;
    practiceTrialData.rt = rt;
    practiceTrialData.response = responseNum;          // 超时为空值 null，否则为标签槽编号（1..nItems）
    practiceTrialData.timeout = timedOut;
    practiceTrialData.correctResponse = correctLabel;
    practiceTrialData.changeItem = changeItemIdx;
    practiceTrialData.origColorIdx = activeColors[changeItemIdx];
    practiceTrialData.probeColorIdx = changeColorIdx;
    practiceTrialData.labelMap = JSON.stringify(labelMap);
    practiceTrialData.trialLabels = JSON.stringify(trialLabels.map(i => LABEL_CHARS[i]));
    practiceTrialData.labelChars = LABEL_CHARS.join('');
    practiceTrialData.accuracy = !timedOut && responseNum === correctLabel ? 1 : 0;

    practiceAcc.push(practiceTrialData.accuracy);
    practiceData.push(Object.assign({}, practiceTrialData));

    drawBackground(); drawFixation();
    practiceTrialIdx++;
    setTimeout(() => runPracticeTrial(), 50);
}

function endPractice() {
    let accPct = Math.round(practiceAcc.reduce((a, b) => a + b, 0) / practiceAcc.length * 100);
    experimentPhase = 'practice_end';
    drawBackground();

    ctx.fillStyle = 'rgb(255,255,255)';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = 'bold 36px Inter, sans-serif';
    ctx.fillText('练习结束', canvas.width / 2, canvas.height / 2 - 80);
    ctx.font = '28px Inter, sans-serif';
    ctx.fillText(`正确率：${accPct}%`, canvas.width / 2, canvas.height / 2 - 20);
    ctx.font = '20px Inter, sans-serif';
    ctx.fillStyle = 'rgb(255,235,59)';
    ctx.fillText('按 空格键(Space) 开始正式实验', canvas.width / 2, canvas.height / 2 + 60);
}

// ============================================================
//  正式实验阶段
// ============================================================
// 生成完整正式试次序列：12 种条件（3 set size × 4 呈现时长）各 30 次，整体只洗牌一次。
// 关键不变量（与 exp1_fcdt 一致）：函数开头把 PRNG 重置回 formalPrngBaseSeed，
// 因此"跳过练习直接续做"与"正常做完练习再进正式"得到的 360 试次序列完全相同。
// 若正式基准起点尚未捕获（formalPrngBaseSeed 为 0），退回用 p.rndSeed 作为起点；
// p.rndSeed 未设置时（如测试沙箱直接调用 buildFormalTrials）再用 Date.now() 兜底，
// 避免 seedRandom(undefined) 污染 PRNG 状态导致 shuffle 无位移。
function buildFormalTrials() {
    const baseSeed = formalPrngBaseSeed !== 0 ? formalPrngBaseSeed : (p.rndSeed || Date.now());
    seedRandom(baseSeed);
    const trials = [];
    for (let ss of prefs.setSizes) {
        for (let duration of prefs.studyDurations) {
            for (let i = 0; i < prefs.nEachCondition; i++) {
                trials.push({ setSize: ss, studyDuration: duration, change: 1 });
            }
        }
    }
    shuffleArray(trials);
    return trials;
}

// 生成后自检：总数、各条件计数和按 50 个切分的 block 结构都必须符合设计。
function verifyFormalTrials(trials) {
    const expectedTotal = prefs.setSizes.length * prefs.studyDurations.length * prefs.nEachCondition;
    const conditionCounts = new Map();
    for (let trial of trials) {
        const key = trial.setSize + '|' + trial.studyDuration;
        conditionCounts.set(key, (conditionCounts.get(key) || 0) + 1);
    }

    let conditionsValid = true;
    for (let ss of prefs.setSizes) {
        for (let duration of prefs.studyDurations) {
            const key = ss + '|' + duration;
            if (conditionCounts.get(key) !== prefs.nEachCondition) conditionsValid = false;
        }
    }

    const blocks = [];
    for (let start = 0; start < trials.length; start += prefs.numTrials) {
        blocks.push(trials.slice(start, start + prefs.numTrials));
    }
    const blocksValid = blocks.length === prefs.numBlocks &&
        blocks.every(block => block.length === prefs.numTrials);
    const valid = trials.length === expectedTotal && conditionsValid && blocksValid;

    if (!valid) {
        console.error('[exp1_clt] 正式试次生成自检未通过（不应发生）：', {
            total: trials.length,
            expectedTotal: expectedTotal,
            conditionCounts: Object.fromEntries(conditionCounts),
            blockSizes: blocks.map(block => block.length),
            expectedBlocks: prefs.numBlocks,
            expectedTrialsPerBlock: prefs.numTrials
        });
    } else {
        console.log('[exp1_clt] 正式试次生成自检通过：' +
            (prefs.setSizes.length * prefs.studyDurations.length) + ' 种条件各 ' + prefs.nEachCondition +
            ' 次，共 ' + trials.length + ' 次，' +
            prefs.numBlocks + ' 个 block × ' + prefs.numTrials + ' 次');
    }
    return valid;
}

function startFormal() {
    allFormalTrials = buildFormalTrials();
    verifyFormalTrials(allFormalTrials);

    // 续做：跳过练习、从上次完成后的下一组开始；否则正常从第 1 组开始。
    currentBlock = Math.min(resumeStartBlock, runtimeNumBlocks);
    if (currentBlock > 1) {
        console.log('[' + EXPERIMENT_NAME + '] 从第 ' + currentBlock + ' 组续做（上次已完成 ' + (currentBlock - 1) + ' 组）');
    }
    startBlock();
}

let allFormalTrials = [];
let formalTrialGlobalIdx = 0;

function startBlock() {
    blockTrials = allFormalTrials.slice(
        (currentBlock - 1) * runtimeNumTrials,
        currentBlock * runtimeNumTrials
    );
    currentTrial = 0;
    formalTrialGlobalIdx = (currentBlock - 1) * runtimeNumTrials;

    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
    switchScreen('experiment');
    runTrial();
}

async function runTrial() {
    if (currentTrial >= blockTrials.length) {
        endBlock();
        return;
    }

    experimentPhase = 'trial_running';
    let condition = blockTrials[currentTrial];
    currentNItems = condition.setSize;
    trialData = {
        subjectID: p.subjectID,
        experimentVersion: EXPERIMENT_VERSION,
        sessionID: SESSION_ID,
        subName: p.subName,
        block: currentBlock,
        trial: currentTrial + 1,
        setSize: condition.setSize,
        studyDuration: condition.studyDuration,
        isChange: 1,
        experimentType: EXPERIMENT_NAME
    };

    // ITI
    drawBackground(); drawFixation();
    await wait(prefs.ITI);

    await showStudyArray(condition);
}

async function showStudyArray(condition) {
    // 一次算好并冻结本试次的视觉计划：记忆阵列 / 探测阵列 / 位置 / 标签
    // 计划内已保证「有且只有一个方块变色」且「新颜色不在记忆阵列中出现」
    const plan = buildProbePlan(condition.setSize);
    trialPlan = plan;
    applyPlanToGlobals(plan);

    trialData.itemLocs = JSON.stringify(plan.locs.map(l => [Math.round(l.x), Math.round(l.y)]));
    trialData.itemColors = JSON.stringify(plan.memColors);

    // 学习阵列（仅方块，无标签）
    drawBackground(); drawFixation();
    drawPlanSquares(plan, plan.memColors);
    const studyStartTime = performance.now();
    await wait(condition.studyDuration);
    trialData.studyDurationActual = (performance.now() - studyStartTime).toFixed(2);

    // 保留间隔（注视点全程固定显示，不消失）
    drawBackground(); drawFixation();
    await wait(prefs.retentionInterval);

    trialData.changeItem = plan.changeIdx;
    trialData.origColorIdx = plan.origColor;
    trialData.probeColorIdx = plan.newColor;
    trialData.labelMap = JSON.stringify(plan.labelMap);
    trialData.trialLabels = JSON.stringify(plan.trialLabels.map(i => LABEL_CHARS[i]));
    trialData.labelChars = LABEL_CHARS.join('');
    trialData.correctResponse = plan.correctLabel;

    // 探测阵列：直接用同一个 plan 的 probeColors，任何情况下都只有一个方块变色
    showTestArray(plan);
}

function showTestArray(plan) {
    // 冻结本试次探测屏内容，供 mousemove 整屏重绘时复用
    probeSnapshot = { locs: plan.locs, colorIdx: plan.probeColors };
    clickedItemIndexCircle = -1;
    resetVirtualCursor();          // 自绘光标回到注视点
    redrawProbeScreen();          // 背景 + 注视点 + 色块 + 自绘光标

    experimentPhase = 'trial_response';
    responseStartTime = performance.now();
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = setTimeout(() => {
        if (experimentPhase === 'trial_response') recordResponse(null);
    }, RESPONSE_LIMIT_MS);
}

function recordResponse(responseNum) {
    // 立即锁定反应阶段：防止鼠标连击在同一试次内触发第二次响应
    if (experimentPhase !== 'trial_response') return;
    experimentPhase = 'trial_locked';
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = null;
    hideVirtualCursor();          // 关掉自绘光标、恢复真实光标

    const timedOut = responseNum === null;
    let rt = timedOut ? null : (performance.now() - responseStartTime) / 1000;
    trialData.rt = rt;
    trialData.response = responseNum;                  // 超时为空值 null，否则为标签槽编号（1..nItems）
    trialData.timeout = timedOut;
    trialData.accuracy = !timedOut && responseNum === correctLabel ? 1 : 0;

    // 记录被试选择颜色与正确颜色的 RGB（色块无固定数字标签，以颜色为准）
    const clickedItemIdx = clickedItemIndexCircle;
    const clickedColorIdx = (clickedItemIdx >= 0 && clickedItemIdx < probeColorArr.length)
        ? probeColorArr[clickedItemIdx]
        : null;
    const correctColorIdx = changeColorIdx;
    trialData.responseColorIdx = clickedColorIdx;
    trialData.correctColorIdx = correctColorIdx;
    trialData.responseRGB = clickedColorIdx !== null ? colors_9[clickedColorIdx].join(',') : null;
    trialData.correctRGB = correctColorIdx !== null ? colors_9[correctColorIdx].join(',') : null;

    // 超时试次标记（append-only 两列，不改动既有列语义）：
    //   TrialStatus = 'valid'   | 'timeout'  —— 人眼在 Trial Data 里可直接按此列筛出超时行
    //   CountedInK  = 1         | 0          —— 0 表示该行不参与 K 值计算（见 buildKRow）
    // 说明：超时行在 Response / RT 两列本来就是空值，但空单元格在 Excel 里容易被当成
    //       "漏填"而与真正的缺失混淆，因此额外写一个显式状态列做区分标记。
    trialData.TrialStatus = timedOut ? 'timeout' : 'valid';
    trialData.CountedInK = timedOut ? 0 : 1;

    // circle 版：记录鼠标动作，供事后核验「各试次移动距离是否一致」。
    // 超时试次留空（不写 0），与项目「缺失留空」约定一致。
    trialData.clickedItemIndex = timedOut ? null : clickedItemIndexCircle;
    trialData.cursorStartX = cursorStartX;
    trialData.cursorStartY = cursorStartY;
    trialData.cursorEndX = timedOut ? null : virtualCursor.x;
    trialData.cursorEndY = timedOut ? null : virtualCursor.y;
    trialData.cursorTravelPx = timedOut ? null : Number(cursorTravelPx.toFixed(2));

    const completedTrial = Object.assign({}, trialData);
    stimData.push(completedTrial);
    if (storage) {
        storage.saveTrial(prepareStorageRow(completedTrial)).catch(error => {
            console.error('[CLT storage] trial save failed:', error);
        });
    }

    drawBackground(); drawFixation();
    currentTrial++;
    formalTrialGlobalIdx++;
    setTimeout(() => runTrial(), 50);
}

// ============================================================
//  Block 间休息（强制 ≥15 秒倒计时，倒计时结束后按 SPACE 进入下一组）
// ============================================================
let breakTimer = null;
const BREAK_MIN_SECONDS = 15;   // block 间强制休息 15 秒

// 计算某个 block 的正确率（%）
function blockAccuracyPct(blockNo) {
    let rows = stimData.filter(d => d.block === blockNo);
    if (rows.length === 0) return 0;
    return Math.round(rows.reduce((s, d) => s + d.accuracy, 0) / rows.length * 100);
}

// 低正确率判定（决定文件名是否加 F，也在 K 值表中标红对应行）：
// 对 set size = 3 和 4 分别统计「呈现时长 ≥ 200 ms」试次的正确率，
// 任一个低于 50% 即触发；set size = 6 不参与判定。
// 返回 { triggered, details: [{ setSize, accuracyPct, low, nTotal }] }；
// 数据不足时 triggered 为 false（details 中 nTotal 记 0），调用方据此不触发
const LOW_ACC_MIN_DURATION_MS = 200;   // 低正确率判定只看呈现时长 ≥ 200 ms 的试次
const LOW_ACC_MIN_SET_SIZE = 4;        // 判定只覆盖 set size 3 和 4（不超过 4）
const LOW_ACC_THRESHOLD_PCT = 50;      // 正确率低于 50% 触发

function computeLowAccuracyFlag() {
    const details = [];
    for (const ss of [3, 4]) {
        const rows = stimData.filter(
            d => d.setSize === ss && Number(d.studyDuration) >= LOW_ACC_MIN_DURATION_MS
        );
        const n = rows.length;
        const accPct = n > 0 ? rows.reduce((s, d) => s + d.accuracy, 0) / n * 100 : 0;
        details.push({
            setSize: ss,
            accuracyPct: n > 0 ? accPct : null,
            nTotal: n,
            low: n > 0 ? accPct < LOW_ACC_THRESHOLD_PCT : false
        });
    }
    return { triggered: details.some(d => d.low), details: details };
}

function endBlock() {
    const blockAcc = blockAccuracyPct(currentBlock);

    // 只统计正确率低于 50% 的 block 数量（写入结果表备查）。
    // 已取消「累计 2 个 block 低于 50% 自动结束实验」的规则：
    // 无论正确率高低，实验都完整跑完预定 block，低正确率只在导出文件名上加 F。
    if (blockAcc < 50) lowAccBlockCount++;

    if (storage) {
        storage.saveCheckpoint({
            subjectID: p.subjectID,
            block: currentBlock,
            blockAccuracy: blockAcc,
            completedTrials: stimData.length,
            nextBlock: currentBlock < runtimeNumBlocks ? currentBlock + 1 : null
        }).catch(error => {
            console.error('[CLT storage] checkpoint save failed:', error);
        });

        // 线上暂存：每完成一个 block 把进度写进 JATOS Batch Session（仅 JATOS 后端实现）。
        // 只存"完成到第几组"；身份哈希由存储层自己注入，这里不传任何个人信息。
        if (currentBlock < runtimeNumBlocks) {
            storage.saveProgress({
                completedBlocks: currentBlock,
                resumeCount: isResumedSession ? resumeCount : 0,
                updatedAt: new Date().toISOString()
            });
        }
    }

    if (currentBlock < runtimeNumBlocks) {
        showBreakScreen(blockAcc);
    } else {
        finishExperiment();
    }
}

function showBreakScreen(blockAcc) {
    switchScreen('break');
    experimentPhase = 'break_countdown';
    document.getElementById('rest-acc').innerText = `您的正确率为 ${blockAcc}%`;
    startBreakCountdown();
}

// 强制休息倒计时（结束后显示等待提示，按 SPACE 进入下一组）
function startBreakCountdown() {
    clearInterval(breakTimer);
    breakTimer = null;

    const cdEl = document.getElementById('rest-countdown');
    const waitEl = document.getElementById('rest-wait');
    let timeLeft = Math.max(BREAK_MIN_SECONDS, Math.round(runtimeBreakLength * 60));

    experimentPhase = 'break_countdown';
    cdEl.innerText = timeLeft;
    cdEl.classList.remove('hidden');
    waitEl.classList.add('hidden');

    breakTimer = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
            clearInterval(breakTimer);
            breakTimer = null;
            cdEl.classList.add('hidden');
            waitEl.classList.remove('hidden');
            experimentPhase = 'break_wait';   // 此时按 SPACE 进入下一组
        } else {
            cdEl.innerText = timeLeft;
        }
    }, 1000);
}

// “继续休息 30 秒”按钮与其 ESC 退出实验功能已删除：
// 倒计时结束后只有「按 SPACE 进入下一组」一种操作。

function nextBlock() {
    clearInterval(breakTimer);
    breakTimer = null;
    currentBlock++;
    startBlock();
}

// 保存前统一补齐被试信息（保证每一行都带完整的被试信息）
function enrichRows(rows) {
    for (let d of rows) {
        Object.assign(d, prepareStorageRow(d));
    }
    return rows;
}

function prepareStorageRow(row) {
    return {
        ...row,
        subGroup: p.subGroup,          // 组别（1/2）：两个范式的先后顺序，导出列名为 Order
        subName: p.subName,
        subGender: p.subGender,
        subAge: p.subAge,
        subIdCard: maskIdCard(p.subIdCard),
        subPhone: p.subPhone,
        subjectID: p.subjectID,
        rndSeed: p.rndSeed,
        // 断点续做标记（append-only，不改现有字段语义；与 exp1_fcdt 一致）：
        //   resumedFromBlock = 本次从第几组续做（0 表示非续做，从头开始）
        //   resumeCount      = 该被试历史上第几次续做（未续做 = 0，第 1 次续做 = 1）
        resumedFromBlock: isResumedSession ? resumedFromBlock : 0,
        resumeCount: isResumedSession ? resumeCount : 0
    };
}

// ============================================================
//  数据保存
// ============================================================
async function finishExperiment() {
    switchScreen('upload');
    experimentPhase = 'upload';

    let uploadMsg = "";
    let uploadFailed = false;

    enrichRows(stimData);
    ss4LowAccuracyFlag = computeLowAccuracyFlag();

    if (storage && stimData.length > 0) {
        try {
            const saveResult = await storage.saveFinalResult({
                trials: stimData,
                probePlanViolations: probePlanViolations,
                lowAccuracyFlag: ss4LowAccuracyFlag,
                lowAccuracyBlockCount: lowAccBlockCount
            });
            uploadMsg += saveResult && saveResult.message ? saveResult.message : '数据保存成功。';
            await storage.finishExperiment({ successful: true, message: 'CLT mouse completed' });
            // 实验正常完成：清掉 Batch Session 里的进度，避免下次误判为"上次没做完"
            await storage.clearProgress();
        } catch (error) {
            console.error('[CLT storage] final save failed:', error);
            uploadFailed = true;
            uploadMsg += '服务器上传失败: ' + error.message + '。';
            try {
                await storage.finishExperiment({ successful: false, message: 'CLT mouse persistence failed' });
            } catch (finishError) {
                console.error('[CLT storage] failed to close unsuccessful study:', finishError);
            }
        }
    } else if (!storage) {
        uploadMsg += "本版本未配置服务器上传，仅生成本地 Excel。";
    }

    if (p.saveLocal && stimData.length > 0) {
        try {
            downloadXLSX();
            uploadMsg += "本地 Excel 已生成下载。";
        } catch (e) {
            console.error('[exp1_clt] 本地导出失败:', e);
            uploadFailed = true;
            uploadMsg += "本地 Excel 导出失败：" + e.message + "。";
        }
    }

    document.getElementById('upload-status').innerText = uploadMsg;

    // 结束语画面保持干净：仅在异常时显示状态说明
    const statusEl = document.getElementById('finished-text');
    if (statusEl) {
        if (uploadFailed) {
            statusEl.innerText = uploadMsg;
            statusEl.classList.remove('hidden');
        } else {
            statusEl.innerText = '';
            statusEl.classList.add('hidden');
        }
    }

    setTimeout(() => { switchScreen('finished'); }, 2000);

    try {
        if (document.exitPointerLock && document.pointerLockElement) {
            document.exitPointerLock();
        }
        if (document.exitFullscreen && document.fullscreenElement) {
            document.exitFullscreen().catch(() => {});
        }
    } catch (e) {}
}

// ============================================================
//  给 K 值表的「低正确率条件行」标红
//  逐单元格写 style（font color: #C00000 深红），避免整行高亮被
//  部分查看器误判为「整行脏数据」。
//
//  当前 vendor 内嵌的 SheetJS 0.18.5 不支持单元格样式写入，
//  此函数在 0.18.5 下是 no-op（函数开头的 canStyle 检查会直接 return）。
//  K 值表里仍写有「LowAccRed」列记录判定结果（1 = 应标红），
//  将来升级 XLSX 到 0.19+ 或改用 ExcelJS 时本函数会自动生效，
//  调用方（downloadXLSX 内的一行调用）无需改动。
// ============================================================
function markKRowsRed(sheet) {
    // 0.18.5 没有样式能力：不写 cell.s，避免污染数据
    const canStyle = false; // XLSX.utils.sheetStyles 在 0.18.5 下 undefined
    if (!canStyle) return;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    // 找 LowAccRed 列与 K 列的相对索引
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: c })];
        headers[c] = cell && cell.v !== undefined ? String(cell.v) : '';
    }
    const redCol = headers.indexOf('LowAccRed');
    const kCol = headers.indexOf('K_Exclude');
    if (redCol === -1) return;

    for (let r = 1; r <= range.e.r; r++) {
        const flagCell = sheet[XLSX.utils.encode_cell({ r: r, c: redCol })];
        const isLow = flagCell && Number(flagCell.v) === 1;
        if (!isLow) continue;
        // K 值为负的条件行排除（避免把 50ms 拉低正确率导致的负 K 误判为「低正确率」）
        if (kCol !== -1) {
            const kCell = sheet[XLSX.utils.encode_cell({ r: r, c: kCol })];
            if (kCell && Number(kCell.v) < 0) continue;
        }
        // 整行标红
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r: r, c: c });
            if (sheet[addr] && sheet[addr].t !== 'z') {
                sheet[addr].s = { font: { color: { rgb: 'C00000' } } };
            }
        }
    }
}

// ============================================================
//  统计一组试次的结果行：两种超时口径的正确率与容量 K
//  durationLabel：该组对应的呈现时长（ms）；传「汇总」表示跨全部时长
//  K = (Acc * N^2 - N) / (N - 1)，Acc 为 0–1 小数
//
//  Sheet2 同时输出两种预先约定的口径（列名 2026-10-22 起统一为短后缀）：
//    Accuracy_Exclude / K_Exclude        ：超时从分母和正确数中剔除（主口径）；
//    Accuracy_TimeoutErr / K_TimeoutErr  ：超时保留在分母，按错误计入（备选口径）。
//  Trial Data 仍用 TrialStatus 与 CountedInK 明确标记超时，TotalTrials、
//  TimeoutTrials、ValidTrials 三列用于核对两种口径的样本数。
// ============================================================
const EXCLUDE_TIMEOUT_FROM_K = true;

function isCountedTrial(d) {
    // 缺列的历史数据（老版本导出的行）按"计入"处理，保证兼容
    if (EXCLUDE_TIMEOUT_FROM_K && (d.CountedInK === 0 || d.timeout === true)) return false;
    return true;
}

function buildKRow(ss, durationLabel, rows) {
    const validRows = rows.filter(isCountedTrial);
    const nValid = validRows.length;
    const nCorrect = validRows.filter(d => d.accuracy === 1).length;
    const accExclude = nValid > 0 ? nCorrect / nValid : 0;
    const nTotal = rows.length;
    const accTimeoutError = nTotal > 0 ? nCorrect / nTotal : 0;

    const correctRTs = validRows.filter(d => d.accuracy === 1 && typeof d.rt === 'number').map(d => d.rt);
    const rtMean = correctRTs.length > 0
        ? correctRTs.reduce((s, v) => s + v, 0) / correctRTs.length
        : '';

    const N = ss;
    const kExclude = N > 1 ? (accExclude * N * N - N) / (N - 1) : '';
    const kTimeoutError = N > 1 ? (accTimeoutError * N * N - N) / (N - 1) : '';

    // 低正确率标红：仅针对「set size ≤ 4 且 呈现时长 ≥ 200 ms」的明细条件行，
    // 且该 set size 在低正确率判定中确实触发了（≥200ms 试次正确率 < 50%）。
    // 汇总行（durationLabel === '汇总'）与 K 为负的条件行都不标红，避免把时长效应混入。
    const isDetailRow = durationLabel !== '汇总' && Number(durationLabel) >= LOW_ACC_MIN_DURATION_MS;
    const setSizeInRule = ss <= LOW_ACC_MIN_SET_SIZE;
    const ssDetail = lowAccDetail.details.find(d => d.setSize === ss);
    const lowAccRed = !!(isDetailRow && setSizeInRule && ssDetail && ssDetail.low && ss4LowAccuracyFlag);

    // 跨时长汇总行不需要 StudyDuration（该列在 K Values 中必须恒为数值）
    const isSummary = String(durationLabel) === '汇总';

    return {
        SubjectID: p.subjectID,
        SessionID: SESSION_ID,
        ExperimentVersion: EXPERIMENT_VERSION,                // 文本：保留前导 0
        ExperimentType: EXPERIMENT_NAME,
        Order: exportNum(p.subGroup),          // 组别 1 / 2（被试间变量）
        SetSize: ss,
        // 汇总行的 StudyDuration 记为 '汇总'（非数值），导出时会被自动留空，
        // 以保证 K Values 的 StudyDuration 列是纯数值自变量（见「导出列规范」）
        StudyDuration: durationLabel,
        Condition: isSummary ? ('SS' + ss + '-汇总') : conditionLabel(durationLabel, ss),
        TotalTrials: nTotal,                   // 该条件全部试次（含超时）
        ValidTrials: nValid,                   // 剔除超时后真正参与统计的试次
        TimeoutTrials: nTotal - nValid,
        TimeoutExclusionRate: (nTotal > 0 ? (nTotal - nValid) / nTotal : 0).toFixed(4),
        CorrectTrials: nCorrect,
        Accuracy_Exclude: accExclude.toFixed(4),        // 主口径：超时剔除
        K_Exclude: typeof kExclude === 'number' ? kExclude.toFixed(4) : '',
        Accuracy_TimeoutErr: accTimeoutError.toFixed(4), // 备选口径：超时按错误计入
        K_TimeoutErr: typeof kTimeoutError === 'number' ? kTimeoutError.toFixed(4) : '',
        MeanRT: rtMean === '' ? '' : rtMean.toFixed(4),  // 只取有效且正确的试次
        LowAccFlag: ss4LowAccuracyFlag ? 1 : 0,
        LowAccRed: lowAccRed ? 1 : 0,
        LowAccBlockCount: lowAccBlockCount
    };
}

// ============================================================
//  按「呈现时长 × set size」分组重排试次
//  分组顺序：呈现时长升序 → set size 升序；组内保持原始实验顺序（不再排序）
//  行首附加 GroupIndex（分组序号，1 起）与 GroupLabel（如 50ms-SS3）两列
//  末尾兜底追加无法归类的行，确保不丢数据
// ============================================================
function groupTrialsByCondition(rows) {
    const durations = prefs.studyDurations.slice().sort((a, b) => a - b);
    const setSizes = prefs.setSizes.slice().sort((a, b) => a - b);

    const grouped = [];
    const placed = new Set();
    let groupIndex = 0;

    for (const duration of durations) {
        for (const ss of setSizes) {
            groupIndex++;
            const groupLabel = duration + 'ms-SS' + ss;
            for (const d of rows) {
                // 用 Number() 比较：字段恒为数字，但避免个别行被写成字符串时漏掉
                if (Number(d.studyDuration) === duration && Number(d.setSize) === ss) {
                    grouped.push(Object.assign({ GroupIndex: groupIndex, GroupLabel: groupLabel }, d));
                    placed.add(d);
                }
            }
        }
    }

    // 兜底：条件无法归类的行（正常不应发生）追加在末尾，避免静默丢数据
    for (const d of rows) {
        if (!placed.has(d)) {
            grouped.push(Object.assign({ GroupIndex: 0, GroupLabel: '未归类' }, d));
        }
    }
    return grouped;
}

// ============================================================
//  导出列规范（2026-10-22，为 R / SPSS 的 GLM / 混合模型而设）
//  ------------------------------------------------------------
//  四条硬规则——以后增删列时必须同步维护，否则统计端会踩坑：
//   ① 长格式（long format）：试次表一行一试次、条件表一行一条件。
//      R 的 lmer/aov、SPSS 的 MIXED 直接吃；不要再手工转宽表。
//   ② 自变量纯数值：SetSize / StudyDuration / Accuracy 等一律写数字；
//      文本只保留在 SubjectID / Condition / TrialStatus 这类标签列。
//      （数值列里混进一个 '汇总'，R 会把整列读成 character。）
//   ③ 缺失留空而不是写 0：超时试次的 RT / Response 写成空单元格，
//      R 读入为 NA、SPSS 读入为缺失值，不会被当成 0 秒 / 0 号反应。
//   ④ 列名纯 ASCII、无空格与中文：避免 R 的 make.names() 自动改名
//      导致脚本里的列名与 Excel 里看到的对不上。
//
//  列顺序（试次表）：标识 → 设计自变量 → 因变量 → 试次状态 → 诊断列
//                    → 被试信息 → 刺激细节（分析时可整块忽略）
// ============================================================

// 转成数值；非数值（空值 / '汇总' / 非数字字符串）一律返回 null → 导出为空单元格
function exportNum(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
}

// 条件标签：50ms-SS3 这类人眼可读的分组名（SPSS 里可直接当因子用）
function conditionLabel(duration, setSize) {
    return duration + 'ms-SS' + setSize;
}

// 按固定列规格生成 sheet：列顺序完全由 columns 决定，与对象键顺序无关
function sheetFromColumns(rows, columns) {
    const headers = columns.map(col => col.header);
    const data = rows.map((row, i) => {
        const out = {};
        for (const col of columns) {
            const v = col.get(row, i);
            if (v === null || v === undefined || v === '') continue;  // 空单元格 → NA
            out[col.header] = col.text ? String(v) : v;
        }
        return out;
    });
    // 显式传 header：即使首行某列缺值，列也不会被挤到后面或消失
    return XLSX.utils.json_to_sheet(data, { header: headers });
}

// --- Sheet 1 / Sheet 4（Trial Data、Grouped Trials）共用 ---
const TRIAL_EXPORT_COLUMNS = [
    { header: 'SubjectID',          text: true, get: (d) => d.subjectID },
    { header: 'SessionID',          text: true, get: (d) => d.sessionID || SESSION_ID },
    { header: 'ExperimentVersion',  text: true, get: (d) => d.experimentVersion || EXPERIMENT_VERSION },
    { header: 'ExperimentType',    text: true, get: (d) => d.experimentType || EXPERIMENT_NAME },
    // Order：组别（1 / 2），两个范式的先后顺序，被试间变量。
    // 1 = 先做变化觉察任务（FCDT）再做变化定位任务（CLT）；2 = 顺序相反（含义见 GROUP_LABELS）
    { header: 'subGroup',                      get: (d) => exportNum(d.subGroup !== undefined ? d.subGroup : p.subGroup) },
    { header: 'Block',                         get: (d) => exportNum(d.block) },
    { header: 'TrialInBlock',                  get: (d) => exportNum(d.trial) },
    { header: 'TrialGlobal',                   get: (d) => exportNum(d.TrialGlobal) },
    { header: 'SetSize',                       get: (d) => exportNum(d.setSize) },
    { header: 'StudyDuration',                 get: (d) => exportNum(d.studyDuration) },
    { header: 'Condition',         text: true, get: (d) => conditionLabel(d.studyDuration, d.setSize) },
    { header: 'Accuracy',                      get: (d) => exportNum(d.accuracy) },
    { header: 'RT',                            get: (d) => (typeof d.rt === 'number' ? Number(d.rt.toFixed(4)) : null) },
    { header: 'ResponseRGB',          text: true, get: (d) => d.responseRGB },
    { header: 'CorrectRGB',           text: true, get: (d) => d.correctRGB },
    { header: 'ResponseColorIdx',               get: (d) => exportNum(d.responseColorIdx) },
    { header: 'CorrectColorIdx',                get: (d) => exportNum(d.correctColorIdx) },
    { header: 'TrialStatus',        text: true, get: (d) => d.TrialStatus || (d.timeout ? 'timeout' : 'valid') },
    { header: 'CountedInK',                   get: (d) => exportNum(d.CountedInK) },
    { header: 'StudyDurationActual',           get: (d) => exportNum(d.studyDurationActual) },
    { header: 'SubName',           text: true, get: (d) => d.subName || p.subName },
    { header: 'SubGender',         text: true, get: (d) => d.subGender || p.subGender },
    { header: 'SubAge',                        get: (d) => exportNum(d.subAge !== undefined ? d.subAge : p.subAge) },
    { header: 'SubIdCard',         text: true, get: (d) => d.subIdCard || maskIdCard(p.subIdCard) },
    { header: 'SubPhone',          text: true, get: (d) => d.subPhone || p.subPhone },
    { header: 'RandomSeed',        text: true, get: (d) => d.rndSeed !== undefined ? d.rndSeed : p.rndSeed },
    { header: 'ResumedFromBlock',              get: (d) => exportNum(d.resumedFromBlock) },
    { header: 'ResumeCount',                   get: (d) => exportNum(d.resumeCount) },
    { header: 'ChangeItem',                    get: (d) => exportNum(d.changeItem) },
    { header: 'OrigColorIdx',                  get: (d) => exportNum(d.origColorIdx) },
    { header: 'ProbeColorIdx',                 get: (d) => exportNum(d.probeColorIdx) },
    { header: 'ItemLocs',          text: true, get: (d) => d.itemLocs },
    { header: 'ItemColors',        text: true, get: (d) => d.itemColors },
    { header: 'LabelMap',         text: true, get: (d) => d.labelMap },
    { header: 'TrialLabels',       text: true, get: (d) => d.trialLabels },
    { header: 'LabelChars',        text: true, get: (d) => d.labelChars },
    // --- circle 版追加列（append-only，不改动上方既有列顺序）---
    { header: 'ClickedItemIndex',            get: (d) => exportNum(d.clickedItemIndex) },
    { header: 'CursorStartX',                get: (d) => exportNum(d.cursorStartX) },
    { header: 'CursorStartY',                get: (d) => exportNum(d.cursorStartY) },
    { header: 'CursorEndX',                  get: (d) => exportNum(d.cursorEndX) },
    { header: 'CursorEndY',                  get: (d) => exportNum(d.cursorEndY) },
    { header: 'CursorTravelPx',              get: (d) => exportNum(d.cursorTravelPx) }
];

// --- Sheet 2（K Values）/ Sheet 3（Summary SetSize）共用 ---
// 列名即 buildKRow 返回的字段名，此处只负责「固定顺序 + 统一类型」
const K_EXPORT_COLUMNS = [
    'SubjectID', 'SessionID', 'ExperimentVersion', 'ExperimentType', 'subGroup', 'SetSize', 'StudyDuration', 'Condition',
    'TotalTrials', 'ValidTrials', 'TimeoutTrials', 'TimeoutExclusionRate',
    'CorrectTrials', 'Accuracy_Exclude', 'K_Exclude',
    'Accuracy_TimeoutErr', 'K_TimeoutErr', 'MeanRT',
    'LowAccFlag', 'LowAccRed', 'LowAccBlockCount'
];
const K_TEXT_COLUMNS = new Set(['SubjectID', 'ExperimentType', 'Condition']);

function kSheetFromRows(rows) {
    const data = rows.map(r => {
        const out = {};
        for (const header of K_EXPORT_COLUMNS) {
            const v = r[header];
            if (v === null || v === undefined || v === '') continue;
            if (K_TEXT_COLUMNS.has(header)) { out[header] = String(v); continue; }
            const n = Number(v);
            // 非数值（如汇总行的 StudyDuration='汇总'）自动留空，
            // 保证 StudyDuration 在 K Values 里恒为数值自变量
            if (Number.isFinite(n)) out[header] = n;
        }
        return out;
    });
    return XLSX.utils.json_to_sheet(data, { header: K_EXPORT_COLUMNS });
}

function downloadXLSX() {
    // 计算低正确率标记（SS3/SS4 在 ≥200ms 正确率 < 50%）
    const lowAccResult = computeLowAccuracyFlag();
    ss4LowAccuracyFlag = lowAccResult.triggered;
    lowAccDetail = lowAccResult;

    // 补全被试信息与断点续做标记（幂等：finishExperiment 里可能已补过一次），
    // 保证 Trial Data 每行都带 SubGender / SubAge / RandomSeed 等被试级字段
    enrichRows(stimData);

    // 全局试次序号（1..360）：长格式里用来查顺序效应 / 疲劳效应，
    // 先按原始实验顺序写好，Grouped Trials 重排后仍指向同一试次
    stimData.forEach((d, i) => { d.TrialGlobal = i + 1; });

    // ============================================================
    //  Sheet 1: Trial Data
    //  长格式，一行一试次，共 360 行；列顺序与含义见上方「导出列规范」
    // ============================================================
    const trialSheet = sheetFromColumns(stimData, TRIAL_EXPORT_COLUMNS);

    // ============================================================
    //  Sheet 2: K Values（条件级，12 行 = 4 呈现时长 × 3 set size）
    //  CLT 公式: K = (Acc * N^2 - N) / (N - 1)
    //  同时报告两种口径：
    //    Accuracy_Exclude / K_Exclude        ：超时剔除（主口径）；
    //    Accuracy_TimeoutErr / K_TimeoutErr ：超时按错误计入（备选口径）。
    //  TimeoutExclusionRate 报告每个条件被剔除的超时比例。
    //  StudyDuration 在本表恒为数值；跨时长的汇总行单独放 Sheet 3。
    // ============================================================
    const kData = [];
    const sortedDurations = prefs.studyDurations.slice().sort((a, b) => a - b);
    const sortedSetSizes = prefs.setSizes.slice().sort((a, b) => a - b);
    for (const duration of sortedDurations) {
        for (const ss of sortedSetSizes) {
            const rows = stimData.filter(d =>
                Number(d.setSize) === ss && Number(d.studyDuration) === duration);
            kData.push(buildKRow(ss, duration, rows));
        }
    }
    const kSheet = kSheetFromRows(kData);
    // 标红低正确率条件行（XLSX 0.18.5 无样式能力时为 no-op，见函数注释）
    markKRowsRed(kSheet);

    // ============================================================
    //  Sheet 3: ColorAccuracy（每个颜色在每个 set size 下的反应正确率）
    //  行：颜色 RGB；列：set size；值：该颜色作为正确选项时的正确率
    // ============================================================
    const colorAccData = [];
    for (const ss of sortedSetSizes) {
        const ssRows = stimData.filter(d => Number(d.setSize) === ss);
        for (let c = 0; c < colors_9.length; c++) {
            const colorRows = ssRows.filter(d => Number(d.correctColorIdx) === c);
            const nTotal = colorRows.length;
            const nCorrect = colorRows.filter(d => d.accuracy === 1).length;
            const acc = nTotal > 0 ? nCorrect / nTotal : null;
            colorAccData.push({
                SetSize: ss,
                ColorIdx: c,
                RGB: colors_9[c].join(','),
                TotalTrials: nTotal,
                CorrectTrials: nCorrect,
                Accuracy: acc
            });
        }
    }
    const colorAccSheet = XLSX.utils.json_to_sheet(colorAccData, {
        header: ['SetSize', 'ColorIdx', 'RGB', 'TotalTrials', 'CorrectTrials', 'Accuracy']
    });

    // ============================================================
    //  Sheet 4: Grouped Trials
    //  列与 Sheet1 完全相同（前面多 GroupIndex / GroupLabel 两列标识分组），
    //  但按「呈现时长 → set size」分组重排：同一条件的试次在表中连续排列，
    //  组内仍保持原始实验顺序，不做二次排序。
    // ============================================================
    const groupedColumns = [
        { header: 'GroupIndex', get: (d) => exportNum(d.GroupIndex) },
        { header: 'GroupLabel', text: true, get: (d) => d.GroupLabel }
    ].concat(TRIAL_EXPORT_COLUMNS);
    const groupedSheet = sheetFromColumns(groupTrialsByCondition(stimData), groupedColumns);

    // === 生成 Excel 文件 ===
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, trialSheet, 'Trial Data');
    XLSX.utils.book_append_sheet(wb, kSheet, 'K Values');
    XLSX.utils.book_append_sheet(wb, colorAccSheet, 'ColorAccuracy');
    XLSX.utils.book_append_sheet(wb, groupedSheet, 'Grouped Trials');

    // 文件名规则：实验名称_被试编号（如 exp1_clt_0112345678.xlsx）
    // 存在低正确率条件（SS3/SS4 在 ≥200ms 正确率 < 50%）时在最前面加 F 标记；
    // 调试模式不做该判定，因此永不带 F
    const filePrefix = (!p.debugMode && ss4LowAccuracyFlag) ? 'F' : '';
    const fileName = filePrefix + EXPERIMENT_NAME + '_' + p.subjectID + '.xlsx';
    console.log(`[exp1_clt] 导出模式=${p.debugMode ? '调试' : '正式'} 被试编号=${p.subjectID} 文件名=${fileName}`);
    saveWorkbookFile(wb, fileName);
}

// ============================================================
//  写出 Excel 文件
//  用正确的 xlsx MIME + 明确的文件名触发下载；同时在结束界面保留一个
//  手动下载链接（某些浏览器/内嵌预览不遵循 download 属性，会把文件存成
//  随机名或没有扩展名，看起来「不是 Excel 文件」，此时可用该链接另存）
// ============================================================
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
let lastExport = null;   // { url, fileName, size }

function saveWorkbookFile(wb, fileName) {
    if (typeof XLSX === 'undefined') {
        throw new Error('Excel 导出库 (xlsx) 未加载');
    }
    const bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const u8 = (bytes instanceof Uint8Array) ? bytes : new Uint8Array(bytes);
    // 校验 zip 魔数 'PK'，确保写出的是合法 xlsx（而不是被当成其它格式）
    if (u8[0] !== 0x50 || u8[1] !== 0x4B) {
        throw new Error('生成的 Excel 内容异常');
    }

    const blob = new Blob([u8], { type: XLSX_MIME });
    const url = URL.createObjectURL(blob);
    lastExport = { url: url, fileName: fileName, size: u8.length };

    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.rel = 'noopener';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();

    console.log(`[exp1_clt] 已导出 ${fileName}（${u8.length} 字节，${wb.SheetNames.length} 个 sheet）`);
    showManualDownload(fileName, url);
}

// 结束界面上的手动下载入口（文件名与自动下载一致）
function showManualDownload(fileName, url) {
    const box = document.getElementById('manual-download');
    if (!box) return;
    const link = document.getElementById('manual-download-link');
    if (link) {
        link.href = url;
        link.download = fileName;
        link.innerText = fileName;
    }
    box.classList.remove('hidden');
}

// ============================================================
//  Canvas 绘图工具
// ============================================================
function drawBackground() {
    ctx.fillStyle = 'rgb(128,128,128)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
}

function drawFixation() {
    ctx.fillStyle = 'rgb(0,0,0)';
    ctx.beginPath();
    ctx.arc(canvas.width / 2, canvas.height / 2, prefs.fixationSize, 0, Math.PI * 2);
    ctx.fill();
}

function drawSquare(x, y, colorArr) {
    ctx.fillStyle = `rgb(${colorArr[0]},${colorArr[1]},${colorArr[2]})`;
    const half = prefs.stimSize / 2;
    ctx.fillRect(x - half, y - half, prefs.stimSize, prefs.stimSize);
}

// ============================================================
//  自绘虚拟光标（circle 版，2026-11-19）
//
//  目的：浏览器不允许网页把「系统光标」放到指定坐标，所以改为
//    1) 探测屏出现时把自绘光标重置到注视点 (canvas 中心)；
//    2) 隐藏真实光标（canvas 上 cursor:none），被试只看得到自绘光标；
//    3) 用 mousemove 的相对位移 movementX/movementY 驱动自绘光标；
//    4) 命中判定一律用自绘光标坐标，不再读 e.clientX/clientY。
//  配合「所有色块中心到注视点距离恒为 R」的圆形布局，
//  无论被试点哪个色块，需要移动的直线距离都相同（= R），
//  从而消除「移动鼠标距离不同」对反应的干扰。
//
//  已知局限（先尝试阶段的预期内问题）：
//    由于已在开始实验时启用 Pointer Lock，真实光标到达屏幕边缘不会
//    让 movementX/movementY 归零，虚拟光标可以覆盖整个画布。
// ============================================================
function drawVirtualCursor() {
    if (!cursorActive) return;
    // 虚拟坐标不设边界；仅将可见标记限制在画布内缘，防止越界后被 Canvas 裁切。
    // 命中判定仍使用未裁切的 virtualCursor 坐标，反向移动回画布后会继续同步。
    const r = 9;
    const pad = r + 2;
    const x = Math.max(pad, Math.min(canvas.width - pad, virtualCursor.x));
    const y = Math.max(pad, Math.min(canvas.height - pad, virtualCursor.y));
    ctx.save();
    // 双色描边：白外圈 + 黑内圈，保证在灰底 (128) 与任意彩色色块上都可见
    ctx.beginPath();
    ctx.arc(x, y, r, 0, Math.PI * 2);
    ctx.lineWidth = 3.5;
    ctx.strokeStyle = 'rgb(255,255,255)';
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(x, y, r, 0, Math.PI * 2);
    ctx.lineWidth = 1.5;
    ctx.strokeStyle = 'rgb(0,0,0)';
    ctx.stroke();
    // 不再绘制十字/短刻度，避免向上移动时出现类似“可关闭叉号”的视觉干扰。
    // 仅保留圆环，中心注视点保持清晰可见。
    ctx.restore();
}

// 重绘探测屏：每次 mousemove 都整屏重画，擦掉上一帧的自绘光标，避免拖影。
// 注视点仍然全程连续显示（每帧先 drawBackground 再 drawFixation）。
function redrawProbeScreen() {
    if (!probeSnapshot) return;
    drawBackground();
    drawFixation();
    for (let i = 0; i < probeSnapshot.locs.length; i++) {
        drawSquare(probeSnapshot.locs[i].x, probeSnapshot.locs[i].y,
                   colors_9[probeSnapshot.colorIdx[i]]);
    }
    drawVirtualCursor();
}

// 每个试次的探测屏开始时调用：自绘光标回到注视点，本试次轨迹归零。
function resetVirtualCursor() {
    virtualCursor.x = canvas.width / 2;
    virtualCursor.y = canvas.height / 2;
    cursorStartX = virtualCursor.x;
    cursorStartY = virtualCursor.y;
    cursorTravelPx = 0;
    lastClientX = null;
    lastClientY = null;
    cursorSkipNextMove = false;  // 每次探测屏重置起点，不丢弃被试的首次真实移动
    cursorActive = true;
    canvas.style.cursor = 'none';   // 隐藏真实光标
}

// 反应之后调用：关掉自绘光标，恢复真实光标（休息 / 结束页需要操作界面）。
function hideVirtualCursor() {
    cursorActive = false;
    probeSnapshot = null;
    canvas.style.cursor = 'default';
}

// 命中判定：返回 (x, y) 落在哪个色块上；未命中返回 -1。
// 与旧版唯一区别是坐标来源改为自绘虚拟光标，而非浏览器事件的 clientX/Y。
function hitTestSquare(x, y) {
    if (!probeSnapshot) return -1;
    const half = prefs.stimSize / 2;
    for (let i = 0; i < probeSnapshot.locs.length; i++) {
        const loc = probeSnapshot.locs[i];
        if (Math.abs(x - loc.x) <= half && Math.abs(y - loc.y) <= half) return i;
    }
    return -1;
}

window.addEventListener('mousemove', (e) => {
    if (!cursorActive) return;
    let dx = 0, dy = 0;
    if (pointerLocked && typeof e.movementX === 'number' && typeof e.movementY === 'number') {
        // 只有确认已锁定指针时才使用 movementX/Y；否则部分浏览器会返回不完整位移。
        dx = e.movementX; dy = e.movementY;
    } else if (lastClientX !== null && Number.isFinite(e.clientX) && Number.isFinite(e.clientY)) {
        // 未锁定时必须使用绝对 clientX/Y 增量，保证上下左右都能移动。
        dx = e.clientX - lastClientX; dy = e.clientY - lastClientY;
    }
    lastClientX = e.clientX;
    lastClientY = e.clientY;
    if (cursorSkipNextMove) { cursorSkipNextMove = false; return; }
    if (dx === 0 && dy === 0) return;

    // 若 canvas 被 CSS 缩放过，把屏幕位移换算回画布坐标
    const rect = canvas.getBoundingClientRect();
    const sx = rect.width ? canvas.width / rect.width : 1;
    const sy = rect.height ? canvas.height / rect.height : 1;
    const mx = dx * sx, my = dy * sy;

    // 探测阶段不限制虚拟光标范围；每次新探测屏开始时才重置到注视点。
    virtualCursor.x += mx;
    virtualCursor.y += my;
    cursorTravelPx += Math.hypot(mx, my);
    redrawProbeScreen();
});

// 指针锁定状态：探测阶段使用锁定后的相对位移驱动虚拟光标。
document.addEventListener('pointerlockchange', () => {
    pointerLocked = (document.pointerLockElement === canvas);
    lastClientX = null;
    lastClientY = null;
    if (!pointerLocked && cursorActive &&
        (experimentPhase === 'trial_response' || experimentPhase === 'practice_response')) {
        hideVirtualCursor();
        alert('鼠标指针锁定已退出，当前试次已停止。请重新进入实验。');
    }
});

document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement && cursorActive &&
        (experimentPhase === 'trial_response' || experimentPhase === 'practice_response')) {
        hideVirtualCursor();
        if (document.pointerLockElement === canvas && document.exitPointerLock) {
            document.exitPointerLock();
        }
        alert('全屏已退出，当前试次已停止。请重新进入实验。');
    }
});

// ============================================================
//  getStimLocs — 单圆 · 四象限布局（circle 版，2026-11-19）
//
//  ① 看不见的参考圆：圆心 = 注视点 (cx, cy)，直径 = 画布短边
//     min(canvas.width, canvas.height)。
//     色块中心落在「该圆内缩半个色块 + 留白」的同心圆上：
//         R = shortSide / 2 - stimSize / 2 - CIRCLE_EDGE_MARGIN
//     内缩是必要的：若把 51px 色块的中心直接放在半径 shortSide/2 上，
//     上下（或左右）边缘的色块会有一半跑出画布，被试看不到完整色块。
//     → 所有色块中心到注视点的距离严格恒为 R。虚拟光标每试次从注视点
//       出发，因此点到任一色块的移动距离完全相同。
//
//  ② 四象限均匀分配：把 setSize 个色块分到 4 个象限，各象限数量尽量均衡
//     （3 → 随机空 1 个象限；4 → 每象限 1 个；6 → 2/2/1/1 系列）。
//     每个半区（左 = 象限 II+III，右 = 象限 IV+I）至少各有 1 个色块，
//     保证 CDT「点左半区色块 / 点右半区色块」两种反应都点得到。
//
//  ③ 角距离随机但均匀：每个象限的可用角区间按该象限的名额等分成若干
//     扇区，在每个扇区的中段随机取一个角度。整体均匀铺开，具体角距随机。
//     区间两端各留 AXIS_MARGIN_DEG，避开 0/90/180/270° 象限坐标轴线，
//     色块既不会压线、也不会横跨左右中线，左右归属无歧义。
//
//  ④ 防重叠：生成多组候选，取「最小两两中心距」最大的一组。
//     半径固定时两点距离 = 2R·sin(Δθ/2)，只需比较角度即可。
// ============================================================
//  getStimLocs — 固定 8 点圆周位置（参考 2021 AP&P）
//
//  预先在注视点周围建立 8 个固定、等角间隔的候选位置（45°）。
//  每个试次只从这 8 个位置中选择 set size 个位置，不再随机生成角度。
//  22.5° 起始角避开水平/垂直坐标轴；候选组合要求左右半区各至少
//  一个位置，并优先保持四象限数量均衡。
// ============================================================
function getStimLocs(nItems, cx, cy) {
    const S = prefs.stimSize;
    const shortSide = Math.min(canvas.width, canvas.height);
    const R = Math.min(
        shortSide * CIRCLE_RADIUS_RATIO,
        shortSide / 2 - S / 2 - CIRCLE_EDGE_MARGIN
    );
    const fixedAngles = Array.from({ length: 8 }, (_, i) =>
        (22.5 + i * 45) * Math.PI / 180
    );
    const candidates = [];
    const choose = (start, picked) => {
        if (picked.length === nItems) {
            const counts = [0, 0, 0, 0];
            for (const idx of picked) counts[Math.floor(idx / 2)]++;
            const left = counts[1] + counts[2];
            const right = counts[0] + counts[3];
            if (left < 1 || right < 1) return;
            candidates.push({ indices: picked.slice(), score: Math.max(...counts) - Math.min(...counts) });
            return;
        }
        for (let i = start; i < 8; i++) {
            picked.push(i); choose(i + 1, picked); picked.pop();
        }
    };
    choose(0, []);
    const bestScore = Math.min(...candidates.map(c => c.score));
    const best = candidates.filter(c => c.score === bestScore);
    const selected = best[randInt(0, best.length - 1)].indices;
    shuffleArray(selected);
    return selected.map(idx => {
        const a = fixedAngles[idx];
        return { x: cx + R * Math.cos(a), y: cy + R * Math.sin(a) };
    });
}
