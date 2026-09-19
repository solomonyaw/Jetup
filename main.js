/**
 * JetUP Investor Information & Account Setup Guide
 * Strictly Pure Vanilla JavaScript Logic
 */
import './style.css';
import * as XLSX from 'xlsx';

// --------------------------------------------------------------------------
// Configuration Constants
// --------------------------------------------------------------------------
export const TAG_MARKETS_URL = 'https://tagmarkets.com';

const APP_CONFIG = {
  registrationUrl: 'https://jetup.ibportal.io/auth/register?e=XiJ1iKWbRG_yO4W1EtstHQcCtTOvJ29__yCIWfHvO6s&a=1',
  tagMarketsUrl: TAG_MARKETS_URL,
  tagMarketsPortalUrl: 'https://user.tagmarkets.com/login',
  jetUpPortalUrl: 'https://jetup.ibportal.io',
  youtubeChannelUrl: 'https://www.youtube.com/@JetUP_official',
  communityTokens: {
    primary: '#9M',
    promo: '9MPromo',
  },
  promoDepositAmount: 150,
  promoBonusAmount: 100,
  promoDisplayedValue: 250,
};

// --------------------------------------------------------------------------
// Calculator State
// --------------------------------------------------------------------------
let calcInputs = {
  initialInvestment: 250,
  dailyReturnPct: 0.5,
  profitSharePct: 20,
  reinvestmentPct: 100,
  numberOfDays: 60,
  additionalDeposit: 0,
};

let showSchedule = false;

// --------------------------------------------------------------------------
// FAQ State & Data (All 20 Questions)
// --------------------------------------------------------------------------
const FAQ_ITEMS = [
  {
    id: 'faq-1',
    category: 'Ecosystem & Roles',
    q: 'What is JetUP?',
    a: 'JetUP is a digital financial ecosystem that brings together trading infrastructure, copy trading, AI-powered trading strategies and a partner/referral program. The ecosystem connects users with TAG Markets and its CopyX platform, where users can access trading strategies such as Sonic AI. JetUP also provides an IB Portal for managing partner-related activities and referral information. Users may participate as traders, copy trading strategies through CopyX, or as partners who refer new participants and potentially earn commissions from qualifying activities within their network.'
  },
  {
    id: 'faq-2',
    category: 'Ecosystem & Roles',
    q: 'Is JetUP a bank or financial institution?',
    a: 'No. JetUP is not a bank, not an asset manager, and not a registered financial adviser. JetUP does not take custody of client funds, does not manage private bank deposits, and does not provide personal investment advice.'
  },
  {
    id: 'faq-3',
    category: 'Ecosystem & Roles',
    q: 'What is TAG Markets?',
    a: 'TAG Markets is an independent online CFD and Forex brokerage firm. It provides the regulated trading infrastructure, trading accounts (MT platforms), execution liquidity, and client margin custody.'
  },
  {
    id: 'faq-4',
    category: 'Ecosystem & Roles',
    q: 'What is CopyX?',
    a: 'CopyX is a specialized copy-trading technology platform connected with TAG Markets. It enables individual client accounts to automatically replicate the trades and positions of verified strategy providers in real time.'
  },
  {
    id: 'faq-5',
    category: 'Ecosystem & Roles',
    q: 'What is Sonic AI?',
    a: 'Sonic AI is an automated copy-trading strategy hosted on the CopyX platform. It utilizes algorithmic parameters primarily focused on gold (XAUUSD) volatility to execute automated trading positions.'
  },
  {
    id: 'faq-6',
    category: 'Custody & Funds',
    q: 'Where are client funds deposited and held?',
    a: 'Client capital is deposited directly into your personal, segregated brokerage account at TAG Markets. JetUP never holds custody of your trading balance, personal bank details, or broker credentials.'
  },
  {
    id: 'faq-7',
    category: 'Custody & Funds',
    q: 'Can I withdraw my money at any time?',
    a: 'Yes, subject to broker clearance and margin requirements. You may request withdrawals of your real deposited capital and realized net profits directly through your TAG Markets client portal at any time.'
  },
  {
    id: 'faq-8',
    category: 'Strategy & Execution',
    q: 'Does Sonic AI trade manually or with algorithms?',
    a: 'Sonic AI operates using automated quantitative algorithms designed to identify price trends and momentum swings in gold (XAUUSD), executing orders programmatically through CopyX.'
  },
  {
    id: 'faq-9',
    category: 'Strategy & Execution',
    q: 'Can copy-trading strategies experience losses or drawdowns?',
    a: 'Yes. All trading involves substantial financial risk. Even advanced automated strategies experience periods of loss and negative drawdowns during adverse market volatility.'
  },
  {
    id: 'faq-10',
    category: 'Promotions',
    q: 'What is the advertised $100 trading bonus?',
    a: 'JetUP highlights a promotion where eligible new users making a qualifying $150 deposit may receive a $100 trading bonus, yielding a $250 displayed starting trading equity.'
  },
  {
    id: 'faq-11',
    category: 'Promotions',
    q: 'Can I withdraw the $100 bonus immediately as cash?',
    a: 'No. The $100 is promotional trading margin credit, not instant withdrawable cash. Specific trading volume requirements, time frames, and broker terms govern bonus eligibility and withdrawal eligibility.'
  },
  {
    id: 'faq-12',
    category: 'Account Setup',
    q: 'Why do I need the #9M Community Token?',
    a: 'The #9M community token authorizes your account connection to the specific partner copy pool on CopyX. Without validating this token, your account will not sync with the master strategy.'
  },
  {
    id: 'faq-13',
    category: 'Account Setup',
    q: 'What is the 9MPromo token used for?',
    a: 'The 9MPromo token is the secondary promotional code entered in CopyX to verify your partner allocation and apply any applicable promotion or community parameters.'
  },
  {
    id: 'faq-14',
    category: 'Ecosystem & Roles',
    q: 'What does "24× Amplify" mean?',
    a: '24× Amplify is an educational concept explaining how leverage and margin amplification work together in CFD trading. It does not represent 24× guaranteed returns.'
  },
  {
    id: 'faq-15',
    category: 'Partner Program',
    q: 'How does the JetUP Partner / Referral Program work?',
    a: 'Participants can introduce new users to JetUP and TAG Markets using personal referral links. Partners may potentially earn introducing broker (IB) commissions based on qualifying trading activity within their network.'
  },
  {
    id: 'faq-16',
    category: 'Partner Program',
    q: 'Are referral earnings guaranteed?',
    a: 'No. Partner commissions depend entirely on actual qualifying trading volume generated by active referred accounts. Commissions are distinct from trading profits and carry no guarantees.'
  },
  {
    id: 'faq-17',
    category: 'Strategy & Execution',
    q: 'What is a profit-share performance fee?',
    a: 'CopyX strategy providers typically deduct an automated performance fee (for example, 20%) strictly from net profitable trades, aligning the strategy provider’s incentives with profitable execution.'
  },
  {
    id: 'faq-18',
    category: 'Risk & Legal',
    q: 'Are trading profits guaranteed?',
    a: 'Absolutely not. No trading system, artificial intelligence model, or automated bot can guarantee returns. Historical performance is never a reliable indicator of future results.'
  },
  {
    id: 'faq-19',
    category: 'Risk & Legal',
    q: 'Can I lose more than my deposit?',
    a: 'Depending on account protections and broker leverage parameters, extreme market slippage can result in total loss of your deposited capital. Always review TAG Markets margin call and stop-out policies.'
  },
  {
    id: 'faq-20',
    category: 'Risk & Legal',
    q: 'What is the Golden Rule of Risk Capital?',
    a: 'Never invest or deposit funds that you cannot afford to lose completely. Only allocate surplus risk capital that will not impact your daily living obligations or financial security if lost.'
  }
];

let activeFaqCategory = 'All';
let faqSearchQuery = '';
let openFaqIds = new Set(['faq-1', 'faq-10', 'faq-12']);

// --------------------------------------------------------------------------
// DOM Initialization
// --------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  initNavbar();
  initHamburgerMenu();
  initCopyButtons();
  initEcosystemTabs();
  initAmplifySlider();
  initCalculator();
  initFaqSection();
  initScrollLinks();
  initTagMarketsButton();
});

// --------------------------------------------------------------------------
// 1. Navbar & Sticky Header
// --------------------------------------------------------------------------
function initNavbar() {
  const navbar = document.getElementById('main-navbar');
  if (!navbar) return;

  window.addEventListener('scroll', () => {
    if (window.scrollY > 20) {
      navbar.classList.add('navbar-scrolled');
    } else {
      navbar.classList.remove('navbar-scrolled');
    }
  }, { passive: true });
}

// --------------------------------------------------------------------------
// 2. Hamburger Menu (Mobile Navigation Dropdown & Dropup)
// --------------------------------------------------------------------------
function initHamburgerMenu() {
  const toggleBtn = document.getElementById('hamburger-toggle-btn');
  const backdrop = document.getElementById('mobile-drawer-backdrop');
  const drawerContent = document.getElementById('mobile-drawer-content');
  const iconBars = document.getElementById('hamburger-icon-bars');
  const iconClose = document.getElementById('hamburger-icon-close');
  const mobileLinks = document.querySelectorAll('.mobile-menu-link, .mobile-menu-cta, #mobile-drawer-content a');

  if (!toggleBtn || !backdrop) return;

  function updatePosition() {
    const navbar = document.getElementById('main-navbar');
    if (navbar && backdrop) {
      const bottom = navbar.getBoundingClientRect().bottom;
      const topOffset = Math.max(0, Math.round(bottom));
      backdrop.style.top = `${topOffset}px`;
      backdrop.style.height = `${window.innerHeight - topOffset}px`;
    }
  }

  function setDrawerOpen(isOpen) {
    toggleBtn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    if (isOpen) {
      updatePosition();
      backdrop.classList.add('active');
      toggleBtn.classList.add('is-active');
      document.body.style.overflow = 'hidden';
      if (iconBars) iconBars.style.display = 'none';
      if (iconClose) iconClose.style.display = 'block';
    } else {
      backdrop.classList.remove('active');
      toggleBtn.classList.remove('is-active');
      document.body.style.overflow = '';
      if (iconBars) iconBars.style.display = 'block';
      if (iconClose) iconClose.style.display = 'none';
    }
  }

  // Toggle button click: opens (dropdown) or closes (dropup)
  toggleBtn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    const isCurrentlyOpen = backdrop.classList.contains('active');
    setDrawerOpen(!isCurrentlyOpen);
  });

  // Close when clicking directly on backdrop overlay outside menu content
  backdrop.addEventListener('click', (e) => {
    if (drawerContent && !drawerContent.contains(e.target)) {
      setDrawerOpen(false);
    }
  });

  // Close when clicking outside anywhere on document
  document.addEventListener('click', (e) => {
    if (backdrop.classList.contains('active')) {
      if (
        drawerContent &&
        !drawerContent.contains(e.target) &&
        !toggleBtn.contains(e.target)
      ) {
        setDrawerOpen(false);
      }
    }
  });

  // Close when clicking any link inside mobile menu & smoothly navigate to section
  mobileLinks.forEach(link => {
    link.addEventListener('click', (e) => {
      const href = link.getAttribute('href');
      setDrawerOpen(false);
      if (href && href.startsWith('#') && href.length > 1) {
        e.preventDefault();
        const targetId = href.substring(1);
        const targetElem = document.getElementById(targetId);
        if (targetElem) {
          setTimeout(() => {
            targetElem.scrollIntoView({ behavior: 'smooth' });
          }, 120);
        }
      }
    });
  });

  // Close on Escape key press
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && backdrop.classList.contains('active')) {
      setDrawerOpen(false);
    }
  });

  // Auto-close on resize to desktop view and sync top position
  window.addEventListener('resize', () => {
    if (window.innerWidth >= 1200) {
      if (backdrop.classList.contains('active')) {
        setDrawerOpen(false);
      }
    } else if (backdrop.classList.contains('active')) {
      updatePosition();
    }
  }, { passive: true });

  window.addEventListener('scroll', () => {
    if (backdrop.classList.contains('active')) {
      updatePosition();
    }
  }, { passive: true });
}

// --------------------------------------------------------------------------
// 3. Clipboard Copy Handlers with Visual Feedback
// --------------------------------------------------------------------------
function initCopyButtons() {
  const copyButtons = document.querySelectorAll('[data-copy-text]');
  copyButtons.forEach(btn => {
    btn.addEventListener('click', async () => {
      const textToCopy = btn.getAttribute('data-copy-text');
      if (!textToCopy) return;

      try {
        await navigator.clipboard.writeText(textToCopy);
        const originalHtml = btn.innerHTML;
        btn.innerHTML = `<span style="color:#10b981;font-weight:bold;">COPIED ✓</span>`;
        setTimeout(() => {
          btn.innerHTML = originalHtml;
        }, 2200);
      } catch (err) {
        // Fallback for older browsers / iframe restrictions
        const textarea = document.createElement('textarea');
        textarea.value = textToCopy;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);

        const originalHtml = btn.innerHTML;
        btn.innerHTML = `<span style="color:#10b981;font-weight:bold;">COPIED ✓</span>`;
        setTimeout(() => {
          btn.innerHTML = originalHtml;
        }, 2200);
      }
    });
  });
}

// --------------------------------------------------------------------------
// 4. Ecosystem Architecture Tabs Toggle
// --------------------------------------------------------------------------
function initEcosystemTabs() {
  const btnTrading = document.getElementById('tab-btn-trading');
  const btnPartner = document.getElementById('tab-btn-partner');
  const panelTrading = document.getElementById('tab-panel-trading');
  const panelPartner = document.getElementById('tab-panel-partner');

  if (!btnTrading || !btnPartner || !panelTrading || !panelPartner) return;

  btnTrading.addEventListener('click', () => {
    btnTrading.classList.add('active');
    btnPartner.classList.remove('active');
    panelTrading.style.display = 'block';
    panelPartner.style.display = 'none';
  });

  btnPartner.addEventListener('click', () => {
    btnPartner.classList.add('active');
    btnTrading.classList.remove('active');
    panelPartner.style.display = 'block';
    panelTrading.style.display = 'none';
  });
}

// --------------------------------------------------------------------------
// 5. 24x Amplify Leverage Visualizer
// --------------------------------------------------------------------------
function initAmplifySlider() {
  const slider = document.getElementById('amplify-slider');
  const displayDeposit = document.getElementById('amplify-val-deposit');
  const displayExposure = document.getElementById('amplify-val-exposure');
  const displayMicroLots = document.getElementById('amplify-val-microlots');
  const displayRiskNotice = document.getElementById('amplify-val-risk-text');

  if (!slider) return;

  function updateAmplifyValues() {
    const deposit = Number(slider.value) || 150;
    const exposure = deposit * 24;
    const microLots = (exposure / 1000).toFixed(1);

    if (displayDeposit) displayDeposit.textContent = `$${deposit.toLocaleString()}`;
    if (displayExposure) displayExposure.textContent = `$${exposure.toLocaleString()}`;
    if (displayMicroLots) displayMicroLots.textContent = `~${microLots} Micro Lots`;
    if (displayRiskNotice) {
      displayRiskNotice.textContent = `A 4.17% adverse market price movement against an unhedged 24× position could result in 100% loss of your $${deposit.toLocaleString()} margin.`;
    }
  }

  slider.addEventListener('input', updateAmplifyValues);
  updateAmplifyValues();
}

// --------------------------------------------------------------------------
// 6. Financial Compounding Calculator & Excel Export
// --------------------------------------------------------------------------
function initCalculator() {
  const initialInvInput = document.getElementById('calc-input-initial');
  const initialInvLabel = document.getElementById('calc-label-initial');
  const dailyReturnSlider = document.getElementById('calc-slider-return');
  const dailyReturnNum = document.getElementById('calc-num-return');
  const dailyReturnLabel = document.getElementById('calc-label-return');
  const profitShareSlider = document.getElementById('calc-slider-share');
  const profitShareLabel = document.getElementById('calc-label-share');
  const reinvestSlider = document.getElementById('calc-slider-reinvest');
  const reinvestLabel = document.getElementById('calc-label-reinvest');
  const daysSlider = document.getElementById('calc-slider-days');
  const daysLabel = document.getElementById('calc-label-days');
  const addDepositInput = document.getElementById('calc-input-add');
  const addDepositLabel = document.getElementById('calc-label-add');
  const resetBtn = document.getElementById('calc-btn-reset');
  const presetButtons = document.querySelectorAll('.calc-preset-btn');
  const scheduleToggleBtn = document.getElementById('calc-toggle-schedule');
  const scheduleWrap = document.getElementById('calc-schedule-wrap');
  const scheduleTbody = document.getElementById('calc-schedule-tbody');
  const downloadExcelBtn = document.getElementById('calc-download-excel-btn');
  const downloadNotice = document.getElementById('calc-download-notice');

  // Outputs
  const outEndingBalance = document.getElementById('calc-out-ending-balance');
  const outNetGain = document.getElementById('calc-out-net-gain');
  const outNetRoi = document.getElementById('calc-out-net-roi');
  const outDayPill = document.getElementById('calc-out-day-pill');
  const outStartCap = document.getElementById('calc-out-start-cap');
  const outGrossProfit = document.getElementById('calc-out-gross-profit');
  const outProfitShare = document.getElementById('calc-out-profit-share');
  const outNetProfit = document.getElementById('calc-out-net-profit');
  const outTotalContrib = document.getElementById('calc-out-total-contrib');
  const outTotalWithdrawn = document.getElementById('calc-out-total-withdrawn');

  function calculateTrajectory() {
    const list = [];
    let currentBalance = calcInputs.initialInvestment;
    let totalContributed = calcInputs.initialInvestment;
    let totalWithdrawn = 0;
    let totalGross = 0;
    let totalShare = 0;
    let totalNet = 0;

    const startDate = new Date();
    const dateOptions = { year: 'numeric', month: 'short', day: 'numeric' };

    for (let day = 1; day <= calcInputs.numberOfDays; day++) {
      const startingBalance = currentBalance;
      const grossProfit = startingBalance * (calcInputs.dailyReturnPct / 100);
      const profitShare = grossProfit * (calcInputs.profitSharePct / 100);
      const netProfit = grossProfit - profitShare;
      const reinvestedProfit = netProfit * (calcInputs.reinvestmentPct / 100);
      const withdrawal = netProfit - reinvestedProfit;
      const additional = calcInputs.additionalDeposit;

      totalGross += grossProfit;
      totalShare += profitShare;
      totalNet += netProfit;
      totalWithdrawn += withdrawal;
      totalContributed += additional;

      const endingBalance = startingBalance + reinvestedProfit + additional;
      currentBalance = endingBalance;

      const dateObj = new Date();
      dateObj.setDate(startDate.getDate() + (day - 1));

      list.push({
        day,
        date: dateObj.toLocaleDateString('en-US', dateOptions),
        startingBalance,
        grossProfit,
        profitShare,
        netProfit,
        reinvestedProfit,
        withdrawal,
        additionalDeposit: additional,
        endingBalance
      });
    }

    const effectiveRoi = totalContributed > 0 ? ((currentBalance + totalWithdrawn - totalContributed) / totalContributed) * 100 : 0;

    return {
      list,
      summary: {
        startingCapital: calcInputs.initialInvestment,
        totalContributions: totalContributed,
        totalWithdrawals: totalWithdrawn,
        estimatedGrossProfit: totalGross,
        profitShareDeducted: totalShare,
        estimatedNetProfit: totalNet,
        projectedEndingBalance: currentBalance,
        effectiveNetROI: effectiveRoi
      }
    };
  }

  function renderCalculator() {
    const { list, summary } = calculateTrajectory();

    // Update Input labels
    if (initialInvLabel) initialInvLabel.textContent = `$${calcInputs.initialInvestment.toLocaleString()}`;
    if (dailyReturnLabel) dailyReturnLabel.textContent = `${calcInputs.dailyReturnPct}%`;
    if (profitShareLabel) profitShareLabel.textContent = `${calcInputs.profitSharePct}%`;
    if (reinvestLabel) reinvestLabel.textContent = `${calcInputs.reinvestmentPct}%`;
    if (daysLabel) daysLabel.textContent = `${calcInputs.numberOfDays} Days`;
    if (addDepositLabel) addDepositLabel.textContent = `$${calcInputs.additionalDeposit}`;

    // Update Output displays
    if (outEndingBalance) {
      outEndingBalance.textContent = `$${summary.projectedEndingBalance.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    if (outNetGain) {
      outNetGain.textContent = `+$${summary.estimatedNetProfit.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    if (outNetRoi) {
      outNetRoi.textContent = `(${summary.effectiveNetROI.toFixed(1)}%)`;
    }
    if (outDayPill) {
      outDayPill.textContent = `Day ${calcInputs.numberOfDays}`;
    }
    if (outStartCap) {
      outStartCap.textContent = `$${summary.startingCapital.toLocaleString()}`;
    }
    if (outGrossProfit) {
      outGrossProfit.textContent = `$${summary.estimatedGrossProfit.toFixed(2)}`;
    }
    if (outProfitShare) {
      outProfitShare.textContent = `-$${summary.profitShareDeducted.toFixed(2)}`;
    }
    if (outNetProfit) {
      outNetProfit.textContent = `$${summary.estimatedNetProfit.toFixed(2)}`;
    }
    if (outTotalContrib) {
      outTotalContrib.textContent = `$${summary.totalContributions.toLocaleString()}`;
    }
    if (outTotalWithdrawn) {
      outTotalWithdrawn.textContent = `$${summary.totalWithdrawals.toFixed(2)}`;
    }

    // Render Table Preview (first 30 rows)
    if (scheduleTbody) {
      const rowsHtml = list.slice(0, 30).map(row => `
        <tr>
          <td style="color:#64748b;font-family:sans-serif;">${row.day} (${row.date})</td>
          <td>$${row.startingBalance.toFixed(2)}</td>
          <td style="color:#0284c7;">+$${row.grossProfit.toFixed(2)}</td>
          <td style="color:#d97706;">-$${row.profitShare.toFixed(2)}</td>
          <td style="color:#059669;">+$${row.netProfit.toFixed(2)}</td>
          <td style="font-weight:700;">$${row.endingBalance.toFixed(2)}</td>
        </tr>
      `).join('');
      scheduleTbody.innerHTML = rowsHtml;
    }

    // Highlight active preset button
    presetButtons.forEach(btn => {
      const val = Number(btn.getAttribute('data-preset-val'));
      if (val === calcInputs.initialInvestment) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });
  }

  // Event Listeners for Inputs
  if (initialInvInput) {
    initialInvInput.addEventListener('input', (e) => {
      calcInputs.initialInvestment = Math.max(0, Number(e.target.value) || 0);
      renderCalculator();
    });
  }

  if (dailyReturnSlider && dailyReturnNum) {
    dailyReturnSlider.addEventListener('input', (e) => {
      calcInputs.dailyReturnPct = Number(e.target.value);
      dailyReturnNum.value = calcInputs.dailyReturnPct;
      renderCalculator();
    });
    dailyReturnNum.addEventListener('input', (e) => {
      calcInputs.dailyReturnPct = Number(e.target.value) || 0;
      dailyReturnSlider.value = calcInputs.dailyReturnPct;
      renderCalculator();
    });
  }

  if (profitShareSlider) {
    profitShareSlider.addEventListener('input', (e) => {
      calcInputs.profitSharePct = Number(e.target.value);
      renderCalculator();
    });
  }

  if (reinvestSlider) {
    reinvestSlider.addEventListener('input', (e) => {
      calcInputs.reinvestmentPct = Number(e.target.value);
      renderCalculator();
    });
  }

  if (daysSlider) {
    daysSlider.addEventListener('input', (e) => {
      calcInputs.numberOfDays = Number(e.target.value);
      renderCalculator();
    });
  }

  if (addDepositInput) {
    addDepositInput.addEventListener('input', (e) => {
      calcInputs.additionalDeposit = Math.max(0, Number(e.target.value) || 0);
      renderCalculator();
    });
  }

  // Preset Buttons
  presetButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const val = Number(btn.getAttribute('data-preset-val'));
      calcInputs.initialInvestment = val;
      if (initialInvInput) initialInvInput.value = val;
      renderCalculator();
    });
  });

  // Reset Button
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      calcInputs = {
        initialInvestment: 250,
        dailyReturnPct: 0.5,
        profitSharePct: 20,
        reinvestmentPct: 100,
        numberOfDays: 60,
        additionalDeposit: 0,
      };
      if (initialInvInput) initialInvInput.value = 250;
      if (dailyReturnSlider) dailyReturnSlider.value = 0.5;
      if (dailyReturnNum) dailyReturnNum.value = 0.5;
      if (profitShareSlider) profitShareSlider.value = 20;
      if (reinvestSlider) reinvestSlider.value = 100;
      if (daysSlider) daysSlider.value = 60;
      if (addDepositInput) addDepositInput.value = 0;
      renderCalculator();
    });
  }

  // Schedule Table Toggle
  if (scheduleToggleBtn && scheduleWrap) {
    scheduleToggleBtn.addEventListener('click', () => {
      showSchedule = !showSchedule;
      scheduleWrap.style.display = showSchedule ? 'block' : 'none';
      scheduleToggleBtn.innerHTML = showSchedule
        ? '<span>Hide Day-by-Day Schedule</span><svg class="icon-inline" viewBox="0 0 24 24"><polyline points="18 15 12 9 6 15"></polyline></svg>'
        : '<span>View Day-by-Day Schedule (Preview)</span><svg class="icon-inline" viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"></polyline></svg>';
    });
  }

  // Excel (.xlsx) Generation & Download
  if (downloadExcelBtn) {
    downloadExcelBtn.addEventListener('click', () => {
      try {
        generateExcelWorkbook(calcInputs);
        if (downloadNotice) {
          downloadNotice.style.display = 'block';
          setTimeout(() => {
            downloadNotice.style.display = 'none';
          }, 4000);
        }
      } catch (err) {
        console.error('Failed to generate Excel file:', err);
      }
    });
  }

  // Initial Calculation Run
  renderCalculator();
}

// --------------------------------------------------------------------------
// 7. Dynamic Excel File Generation (SheetJS / XLSX)
// --------------------------------------------------------------------------
function generateExcelWorkbook(inputs) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Simulation Inputs
  const inputsData = [
    ['JETUP INVESTOR FINANCIAL PLANNING MODEL', ''],
    ['HYPOTHETICAL SIMULATION PARAMETERS', ''],
    ['', ''],
    ['Parameter Name', 'Simulated Value', 'Units / Description'],
    ['Initial Investment Capital', inputs.initialInvestment, 'USD (e.g. $150 deposit + $100 promo bonus)'],
    ['Hypothetical Daily Return', inputs.dailyReturnPct / 100, 'Daily % Assumption (Illustrative)'],
    ['Profit Share Deduction', inputs.profitSharePct / 100, 'Performance Fee % Deducted on Gross'],
    ['Reinvestment / Compounding Rate', inputs.reinvestmentPct / 100, '% of Net Profit Reinvested in Balance'],
    ['Simulation Duration', inputs.numberOfDays, 'Trading Days'],
    ['Additional Periodic Deposit', inputs.additionalDeposit, 'USD added daily'],
    ['', ''],
    ['REGULATORY NOTICE', ''],
    ['This workbook is an educational mathematical projection tool only.', ''],
    ['It does not guarantee trading returns. CFD/Forex trading carries substantial risk of capital loss.', '']
  ];
  const wsInputs = XLSX.utils.aoa_to_sheet(inputsData);
  XLSX.utils.book_append_sheet(wb, wsInputs, 'Inputs');

  // Sheet 2: Day-by-Day Compounding Projection
  const projectionHeaders = [
    'Day',
    'Date',
    'Starting Balance ($)',
    'Gross Profit ($)',
    'Profit Share ($)',
    'Net Profit ($)',
    'Reinvested ($)',
    'Withdrawal ($)',
    'Additional Deposit ($)',
    'Ending Balance ($)'
  ];

  const projectionRows = [projectionHeaders];
  let currentBalance = inputs.initialInvestment;
  const startDate = new Date();

  for (let day = 1; day <= inputs.numberOfDays; day++) {
    const startBal = currentBalance;
    const gross = startBal * (inputs.dailyReturnPct / 100);
    const share = gross * (inputs.profitSharePct / 100);
    const net = gross - share;
    const reinvested = net * (inputs.reinvestmentPct / 100);
    const withdrawal = net - reinvested;
    const additional = inputs.additionalDeposit;
    const ending = startBal + reinvested + additional;
    currentBalance = ending;

    const dateObj = new Date();
    dateObj.setDate(startDate.getDate() + (day - 1));

    projectionRows.push([
      day,
      dateObj.toISOString().split('T')[0],
      Number(startBal.toFixed(2)),
      Number(gross.toFixed(2)),
      Number(share.toFixed(2)),
      Number(net.toFixed(2)),
      Number(reinvested.toFixed(2)),
      Number(withdrawal.toFixed(2)),
      Number(additional.toFixed(2)),
      Number(ending.toFixed(2))
    ]);
  }

  const wsProjection = XLSX.utils.aoa_to_sheet(projectionRows);
  XLSX.utils.book_append_sheet(wb, wsProjection, 'Projection');

  // Sheet 3: Financial Summary & Metrics
  const lastRow = projectionRows[projectionRows.length - 1];
  const endingBal = lastRow[9];
  const totalContributed = inputs.initialInvestment + (inputs.additionalDeposit * inputs.numberOfDays);
  const netEstimatedGain = endingBal - totalContributed;
  const effectiveRoi = totalContributed > 0 ? (netEstimatedGain / totalContributed) * 100 : 0;

  const summaryData = [
    ['SIMULATION RESULTS SUMMARY', ''],
    ['', ''],
    ['Metric Description', 'Value (USD / %)'],
    ['Initial Starting Capital', inputs.initialInvestment],
    ['Total Principal Contributed', totalContributed],
    ['Projected Ending Balance', Number(endingBal.toFixed(2))],
    ['Total Estimated Net Profit', Number(netEstimatedGain.toFixed(2))],
    ['Effective Net ROI %', `${effectiveRoi.toFixed(2)}%`],
    ['Total Trading Days Simulated', inputs.numberOfDays],
    ['', ''],
    ['CRITICAL RISK DISCLAIMER', ''],
    ['Past results and algorithmic projections do not guarantee future profitability.', ''],
    ['Trading Forex and CFDs involves substantial market risk of capital loss.', ''],
    ['Generated by JetUP Independent Investor Guide v2.4', new Date().toISOString()]
  ];

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  // Trigger File Download
  XLSX.writeFile(wb, `JetUP_Compounding_Model_${inputs.initialInvestment}USD_${inputs.numberOfDays}Days.xlsx`);
}

// --------------------------------------------------------------------------
// 8. FAQ Accordion, Category Filters & Search
// --------------------------------------------------------------------------
function initFaqSection() {
  const container = document.getElementById('faq-accordion-container');
  const searchInput = document.getElementById('faq-search-input');
  const catButtons = document.querySelectorAll('.faq-cat-btn');

  if (!container) return;

  function renderFaqItems() {
    const query = faqSearchQuery.toLowerCase().trim();
    const filtered = FAQ_ITEMS.filter(item => {
      const matchesCategory = activeFaqCategory === 'All' || item.category === activeFaqCategory;
      const matchesQuery = !query || item.q.toLowerCase().includes(query) || item.a.toLowerCase().includes(query);
      return matchesCategory && matchesQuery;
    });

    if (filtered.length === 0) {
      container.innerHTML = `
        <div style="text-align:center;padding:3rem 1rem;color:#64748b;background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;">
          <p style="font-size:0.95rem;font-weight:600;">No matching questions found.</p>
          <p style="font-size:0.8rem;margin-top:0.35rem;">Try clearing your search query or selecting "All Categories".</p>
        </div>
      `;
      return;
    }

    container.innerHTML = filtered.map(item => {
      const isOpen = openFaqIds.has(item.id);
      return `
        <div class="faq-accordion-item ${isOpen ? 'open' : ''}" id="faq-item-${item.id}">
          <button class="faq-accordion-trigger" data-faq-id="${item.id}" aria-expanded="${isOpen ? 'true' : 'false'}">
            <span style="font-size:0.925rem;font-weight:700;color:#0f172a;text-align:left;">${item.q}</span>
            <span style="display:flex;align-items:center;color:#7c3aed;flex-shrink:0;">
              <svg class="icon-inline" style="transform:${isOpen ? 'rotate(180deg)' : 'none'};transition:transform 0.2s ease;" viewBox="0 0 24 24">
                <polyline points="6 9 12 15 18 9"></polyline>
              </svg>
            </span>
          </button>
          <div class="faq-accordion-body" style="display:${isOpen ? 'block' : 'none'};">
            <div style="font-size:0.75rem;font-weight:700;color:#7c3aed;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:0.35rem;">
              ${item.category}
            </div>
            <p>${item.a}</p>
          </div>
        </div>
      `;
    }).join('');

    // Attach click triggers
    const triggers = container.querySelectorAll('.faq-accordion-trigger');
    triggers.forEach(trig => {
      trig.addEventListener('click', () => {
        const id = trig.getAttribute('data-faq-id');
        if (openFaqIds.has(id)) {
          openFaqIds.delete(id);
        } else {
          openFaqIds.add(id);
        }
        renderFaqItems();
      });
    });
  }

  // Category Pill Buttons
  catButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      catButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeFaqCategory = btn.getAttribute('data-faq-cat') || 'All';
      renderFaqItems();
    });
  });

  // Search Input
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      faqSearchQuery = e.target.value || '';
      renderFaqItems();
    });
  }

  renderFaqItems();
}

// --------------------------------------------------------------------------
// 9. Smooth Scroll to Top & Anchor Links
// --------------------------------------------------------------------------
function initScrollLinks() {
  const scrollTopBtns = document.querySelectorAll('[data-scroll-top]');
  scrollTopBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });
}

// --------------------------------------------------------------------------
// 10. Learn More About TAG Markets Handler
// --------------------------------------------------------------------------
function initTagMarketsButton() {
  const btn = document.getElementById('btn-learn-tag-markets');
  if (btn) {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      window.open(TAG_MARKETS_URL, '_blank', 'noopener,noreferrer');
    });
  }
}
