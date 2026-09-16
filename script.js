/**
 * Approtech Internship Certificate Form - Interactive Logic
 * Vanilla JavaScript (ES6+)
 */

document.addEventListener('DOMContentLoaded', () => {
  // Global State
  let currentStep = 1;
  const totalSteps = 5;
  const STORAGE_KEY = 'approtech_cert_form_data_2026';

  // DOM Elements - Navigation & Headers
  const form = document.getElementById('certificate-form');
  const btnPrev = document.getElementById('btn-prev');
  const btnNext = document.getElementById('btn-next');
  const btnSubmit = document.getElementById('btn-submit');
  const progressBarFill = document.getElementById('progress-bar-fill');
  const progressRingFill = document.getElementById('progress-ring-fill');
  const ringText = document.getElementById('ring-text');
  const stepCounterText = document.getElementById('step-counter-text');

  // Desktop & Mobile Start Over Buttons
  const desktopStartOverBtn = document.getElementById('desktop-startover-btn');
  const mobileStartOverBtn = document.getElementById('mobile-startover-btn');

  // Modals
  const modalWelcome = document.getElementById('modal-welcome');
  const btnWelcomeResume = document.getElementById('btn-welcome-resume');
  const btnWelcomeReset = document.getElementById('btn-welcome-reset');

  const modalStartOver = document.getElementById('modal-startover');
  const btnStartOverCancel = document.getElementById('btn-startover-cancel');
  const btnStartOverConfirm = document.getElementById('btn-startover-confirm');

  // Form Inputs
  const fields = {
    fullName: document.getElementById('fullName'),
    registerNo: document.getElementById('registerNo'),
    email: document.getElementById('email'),
    phone: document.getElementById('phone'),
    collegeName: document.getElementById('collegeName'),
    degreeBranch: document.getElementById('degreeBranch'),
    state: document.getElementById('state'),
    domain: document.getElementById('domain'),
    startDate: document.getElementById('startDate'),
    endDate: document.getElementById('endDate'),
    projectTitle: document.getElementById('projectTitle')
  };

  const durationVal = document.getElementById('duration-val');
  const durationBox = document.getElementById('duration-box');

  // Success Screen Elements
  const successScreen = document.getElementById('success-screen');
  const successRefId = document.getElementById('success-ref-id');
  const btnCopyRef = document.getElementById('btn-copy-ref');
  const copyText = document.getElementById('copy-text');
  const btnAnotherResponse = document.getElementById('btn-another-response');

  // Certificate Card Details
  const certStudentName = document.getElementById('cert-student-name');
  const certDomainTag = document.getElementById('cert-domain-tag');
  const certModeTag = document.getElementById('cert-mode-tag');
  const certCollege = document.getElementById('cert-college');
  const certRegno = document.getElementById('cert-regno');
  const certDurationVal = document.getElementById('cert-duration-val');
  const certDateStamp = document.getElementById('cert-date-stamp');

  // ==========================================================================
  // 1. INITIALIZATION & STORAGE CHECK
  // ==========================================================================
  function initializeForm() {
    setupInputListeners();
    setupNavigationListeners();
    setupKeyboardListeners();
    checkSavedProgress();
    updateUI();
  }

  function checkSavedProgress() {
    const rawData = localStorage.getItem(STORAGE_KEY);
    if (!rawData) return;

    try {
      const saved = JSON.parse(rawData);
      // Check if there is any meaningful data saved
      const hasData = Object.keys(saved.data || {}).some(k => saved.data[k] && saved.data[k].trim() !== '');
      if (hasData) {
        modalWelcome.classList.remove('hidden');
      }
    } catch (e) {
      console.warn('Could not parse saved progress', e);
    }
  }

  // ==========================================================================
  // 2. INPUT INTERACTION & LIVE VALIDATION
  // ==========================================================================
  function setupInputListeners() {
    // Standard inputs
    Object.keys(fields).forEach(key => {
      const input = fields[key];
      if (!input) return;

      input.addEventListener('input', () => {
        validateField(key);
        if (key === 'startDate' || key === 'endDate') {
          handleDateChange();
        }
        saveProgress();
      });

      input.addEventListener('blur', () => {
        validateField(key);
      });
    });

    // Mode Radio Cards (Online / Offline)
    const modeRadios = document.querySelectorAll('input[name="mode"]');
    modeRadios.forEach(radio => {
      radio.addEventListener('change', () => {
        validateModeField();
        saveProgress();
      });
    });

    // Phone prefix formatting restriction (digits only)
    if (fields.phone) {
      fields.phone.addEventListener('input', (e) => {
        e.target.value = e.target.value.replace(/\D/g, '');
      });
    }
  }

  // Date constraints & dynamic duration calculation
  function handleDateChange() {
    const startVal = fields.startDate.value;
    const endVal = fields.endDate.value;

    if (startVal) {
      fields.endDate.min = startVal;
      if (endVal && endVal < startVal) {
        fields.endDate.value = '';
      }
    }

    calculateDuration();
  }

  function calculateDuration() {
    const startVal = fields.startDate.value;
    const endVal = fields.endDate.value;

    if (!startVal || !endVal) {
      durationVal.textContent = 'Select start and end dates';
      durationBox.style.borderColor = 'var(--accent-cyan)';
      return null;
    }

    const start = new Date(startVal);
    const end = new Date(endVal);

    if (end < start) {
      durationVal.textContent = 'End date cannot be earlier than start date';
      durationBox.style.borderColor = 'var(--error)';
      return null;
    }

    const diffTime = Math.abs(end - start);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // Including start day

    const weeks = Math.round(diffDays / 7);
    let str = `${diffDays} days`;
    if (weeks >= 1) {
      str += ` (~ ${weeks} ${weeks === 1 ? 'week' : 'weeks'})`;
    }

    durationVal.textContent = str;
    durationBox.style.borderColor = 'var(--success)';
    return { days: diffDays, text: str };
  }

  // ==========================================================================
  // 3. SMART VALIDATION LOGIC
  // ==========================================================================
  function getSelectedMode() {
    const selected = document.querySelector('input[name="mode"]:checked');
    return selected ? selected.value : '';
  }

  function validateModeField() {
    const selected = getSelectedMode();
    const errEl = document.getElementById('err-mode');
    if (!selected) {
      errEl.textContent = 'Please select an internship mode (Online or Offline).';
      return false;
    } else {
      errEl.textContent = '';
      return true;
    }
  }

  function validateField(fieldKey) {
    const input = fields[fieldKey];
    if (!input) return true;

    const val = input.value ? input.value.trim() : '';
    const wrapper = input.closest('.input-wrapper');
    const errEl = document.getElementById(`err-${fieldKey}`);
    let isValid = true;
    let errMsg = '';

    switch (fieldKey) {
      case 'fullName':
        if (!val) {
          errMsg = 'Please enter your full name.';
          isValid = false;
        } else if (val.length < 3) {
          errMsg = 'Full name must be at least 3 characters.';
          isValid = false;
        }
        break;

      case 'registerNo':
        if (!val) {
          errMsg = 'Please enter your register / roll number.';
          isValid = false;
        }
        break;

      case 'email':
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!val) {
          errMsg = 'Please enter your email address.';
          isValid = false;
        } else if (!emailRegex.test(val)) {
          errMsg = 'Please enter a valid email format (e.g. name@example.com).';
          isValid = false;
        }
        break;

      case 'phone':
        if (!val) {
          errMsg = 'Please enter your 10-digit mobile number.';
          isValid = false;
        } else if (val.length !== 10 || !/^[6-9]\d{9}$/.test(val)) {
          errMsg = 'Please enter a valid 10-digit Indian phone number.';
          isValid = false;
        }
        break;

      case 'collegeName':
        if (!val) {
          errMsg = 'Please enter your college or institution name.';
          isValid = false;
        }
        break;

      case 'degreeBranch':
        if (!val) {
          errMsg = 'Please enter your degree and branch.';
          isValid = false;
        }
        break;

      case 'state':
        if (!val) {
          errMsg = 'Please select your state.';
          isValid = false;
        }
        break;

      case 'domain':
        if (!val) {
          errMsg = 'Please select your internship domain.';
          isValid = false;
        }
        break;

      case 'startDate':
        if (!val) {
          errMsg = 'Please select your internship start date.';
          isValid = false;
        }
        break;

      case 'endDate':
        if (!val) {
          errMsg = 'Please select your internship end date.';
          isValid = false;
        } else if (fields.startDate.value && val < fields.startDate.value) {
          errMsg = 'End date cannot be earlier than start date.';
          isValid = false;
        }
        break;

      case 'projectTitle':
        if (!val) {
          errMsg = 'Please enter your project title.';
          isValid = false;
        } else if (val.length < 3) {
          errMsg = 'Project title must be at least 3 characters.';
          isValid = false;
        }
        break;

      default:
        break;
    }

    if (errEl) {
      errEl.textContent = errMsg;
    }

    if (wrapper) {
      if (isValid && val) {
        wrapper.classList.add('is-valid');
        wrapper.classList.remove('has-error');
      } else if (!isValid) {
        wrapper.classList.remove('is-valid');
        wrapper.classList.add('has-error');
      } else {
        wrapper.classList.remove('is-valid', 'has-error');
      }
    }

    return isValid;
  }

  function validateStep(stepNum) {
    let isValid = true;

    if (stepNum === 1) {
      const v1 = validateField('fullName');
      const v2 = validateField('registerNo');
      const v3 = validateField('email');
      const v4 = validateField('phone');
      isValid = v1 && v2 && v3 && v4;
    } else if (stepNum === 2) {
      const v1 = validateField('collegeName');
      const v2 = validateField('degreeBranch');
      const v3 = validateField('state');
      isValid = v1 && v2 && v3;
    } else if (stepNum === 3) {
      const v1 = validateField('domain');
      const v2 = validateModeField();
      const v3 = validateField('startDate');
      const v4 = validateField('endDate');
      isValid = v1 && v2 && v3 && v4;
    } else if (stepNum === 4) {
      isValid = validateField('projectTitle');
    }

    return isValid;
  }

  // ==========================================================================
  // 4. STEP NAVIGATION & REVIEW POPULATION
  // ==========================================================================
  function setupNavigationListeners() {
    btnNext.addEventListener('click', () => {
      if (validateStep(currentStep)) {
        goToStep(currentStep + 1);
      }
    });

    btnPrev.addEventListener('click', () => {
      if (currentStep > 1) {
        goToStep(currentStep - 1);
      }
    });

    // Journey Step Click Handler (Desktop)
    document.querySelectorAll('.journey-step-item').forEach(item => {
      item.addEventListener('click', () => {
        const targetStep = parseInt(item.getAttribute('data-journey-step'));
        if (targetStep < currentStep) {
          goToStep(targetStep);
        } else if (targetStep > currentStep) {
          if (validateStep(currentStep)) {
            goToStep(targetStep);
          }
        }
      });
    });

    // Step 5 Edit Buttons
    document.querySelectorAll('.btn-edit-step').forEach(btn => {
      btn.addEventListener('click', () => {
        const targetStep = parseInt(btn.getAttribute('data-edit-target'));
        goToStep(targetStep);
      });
    });

    // Form Submit Event
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      if (validateStep(currentStep)) {
        handleFormSubmission();
      }
    });

    // Start Over Trigger Buttons
    [desktopStartOverBtn, mobileStartOverBtn].forEach(btn => {
      if (btn) {
        btn.addEventListener('click', () => {
          modalStartOver.classList.remove('hidden');
        });
      }
    });

    // Modal Action Listeners
    btnStartOverCancel.addEventListener('click', () => {
      modalStartOver.classList.add('hidden');
    });

    btnStartOverConfirm.addEventListener('click', () => {
      modalStartOver.classList.add('hidden');
      clearProgress();
    });

    btnWelcomeResume.addEventListener('click', () => {
      modalWelcome.classList.add('hidden');
      loadProgress();
    });

    btnWelcomeReset.addEventListener('click', () => {
      modalWelcome.classList.add('hidden');
      clearProgress();
    });

    // Copy Reference ID Button
    btnCopyRef.addEventListener('click', () => {
      const refText = successRefId.textContent;
      navigator.clipboard.writeText(refText).then(() => {
        copyText.textContent = 'Copied!';
        btnCopyRef.style.background = 'var(--success)';
        setTimeout(() => {
          copyText.textContent = 'Copy';
          btnCopyRef.style.background = 'var(--accent)';
        }, 2000);
      });
    });

    // Submit Another Response
    btnAnotherResponse.addEventListener('click', () => {
      clearProgress();
    });
  }

  function goToStep(stepNum) {
    if (stepNum < 1 || stepNum > totalSteps) return;

    // Hide all steps
    document.querySelectorAll('.form-step').forEach(step => {
      step.classList.remove('active');
    });

    currentStep = stepNum;

    // Show current step
    const currentStepEl = document.getElementById(`step-${currentStep}`);
    if (currentStepEl) {
      currentStepEl.classList.add('active');
    }

    // Populate Review Screen if Step 5
    if (currentStep === 5) {
      populateReviewScreen();
    }

    updateUI();
    saveProgress();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  let currentProgressPct = 20;
  let progressAnimId = null;

  function animateProgressRing(targetPct, isCompletedStep) {
    const ringBox = document.getElementById('progress-ring-box');
    const ringFill = document.getElementById('progress-ring-fill');
    const ringText = document.getElementById('ring-text');

    if (!ringBox || !ringFill || !ringText) return;

    // Trigger ring pulse animation
    ringBox.classList.remove('ring-step-pulse');
    void ringBox.offsetWidth; // Force reflow
    ringBox.classList.add('ring-step-pulse');

    if (isCompletedStep) {
      ringBox.classList.add('is-complete');
    } else {
      ringBox.classList.remove('is-complete');
    }

    // SVG dashoffset transition
    const circumference = 144.51; // 2 * PI * 23
    const targetOffset = circumference - (targetPct / 100) * circumference;
    ringFill.style.strokeDashoffset = targetOffset;

    // Animated percentage counter
    if (progressAnimId) cancelAnimationFrame(progressAnimId);

    const startPct = currentProgressPct;
    const duration = 550; // ms matching stroke transition
    const startTime = performance.now();

    function stepCounter(now) {
      const elapsed = now - startTime;
      const progress = Math.min(elapsed / duration, 1);
      // Ease Out Cubic
      const easeProgress = 1 - Math.pow(1 - progress, 3);
      const val = Math.round(startPct + (targetPct - startPct) * easeProgress);

      currentProgressPct = val;

      if (isCompletedStep && progress >= 0.75) {
        ringText.innerHTML = '<i class="fa-solid fa-check check-icon" style="color:var(--success);font-size:1.05rem;"></i>';
      } else {
        ringText.textContent = `${val}%`;
      }

      if (progress < 1) {
        progressAnimId = requestAnimationFrame(stepCounter);
      } else {
        currentProgressPct = targetPct;
        if (isCompletedStep) {
          ringText.innerHTML = '<i class="fa-solid fa-check check-icon" style="color:var(--success);font-size:1.05rem;"></i>';
        } else {
          ringText.textContent = `${targetPct}%`;
        }
      }
    }

    // Trigger text bump animation
    ringText.classList.remove('text-bump');
    void ringText.offsetWidth;
    ringText.classList.add('text-bump');

    progressAnimId = requestAnimationFrame(stepCounter);
  }

  function updateUI() {
    // 1. Update Step Counter & Buttons
    stepCounterText.textContent = `Step 0${currentStep} of 0${totalSteps}`;

    if (currentStep === 1) {
      btnPrev.disabled = true;
    } else {
      btnPrev.disabled = false;
    }

    if (currentStep === totalSteps) {
      btnNext.classList.add('hidden');
      btnSubmit.classList.remove('hidden');
    } else {
      btnNext.classList.remove('hidden');
      btnSubmit.classList.add('hidden');
    }

    // 2. Update Progress Bar
    const progressPct = (currentStep / totalSteps) * 100;
    progressBarFill.style.width = `${progressPct}%`;

    // 3. Update Certificate Progress Ring (SVG dashoffset & smooth animated transition)
    const isCompletedStep = (currentStep === totalSteps);
    animateProgressRing(progressPct, isCompletedStep);

    // 4. Update Left Desktop Journey Panel Steps
    document.querySelectorAll('.journey-step-item').forEach(item => {
      const stepVal = parseInt(item.getAttribute('data-journey-step'));
      item.classList.remove('active', 'completed');
      if (stepVal === currentStep) {
        item.classList.add('active');
      } else if (stepVal < currentStep) {
        item.classList.add('completed');
      }
    });

    // 5. Update Mobile Compact Step Pills
    document.querySelectorAll('.pill-step').forEach(pill => {
      const stepVal = parseInt(pill.getAttribute('data-pill-step'));
      pill.classList.remove('active', 'completed');
      if (stepVal === currentStep) {
        pill.classList.add('active');
      } else if (stepVal < currentStep) {
        pill.classList.add('completed');
      }
    });
  }

  function populateReviewScreen() {
    document.getElementById('rev-fullName').textContent = fields.fullName.value || '-';
    document.getElementById('rev-registerNo').textContent = fields.registerNo.value || '-';
    document.getElementById('rev-email').textContent = fields.email.value || '-';
    document.getElementById('rev-phone').textContent = fields.phone.value ? `+91 ${fields.phone.value}` : '-';

    document.getElementById('rev-collegeName').textContent = fields.collegeName.value || '-';
    document.getElementById('rev-degreeBranch').textContent = fields.degreeBranch.value || '-';
    document.getElementById('rev-state').textContent = fields.state.value || '-';

    document.getElementById('rev-domain').textContent = fields.domain.value || '-';
    document.getElementById('rev-mode').textContent = getSelectedMode() || '-';
    document.getElementById('rev-startDate').textContent = formatDateStr(fields.startDate.value);
    document.getElementById('rev-endDate').textContent = formatDateStr(fields.endDate.value);

    const dur = calculateDuration();
    document.getElementById('rev-duration').textContent = dur ? dur.text : '-';

    document.getElementById('rev-projectTitle').textContent = fields.projectTitle.value || '-';
  }

  function formatDateStr(dateStr) {
    if (!dateStr) return '-';
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
  }

  // ==========================================================================
  // 5. AUTO-SAVE & LOCALSTORAGE LOGIC
  // ==========================================================================
  function saveProgress() {
    const data = {
      fullName: fields.fullName.value,
      registerNo: fields.registerNo.value,
      email: fields.email.value,
      phone: fields.phone.value,
      collegeName: fields.collegeName.value,
      degreeBranch: fields.degreeBranch.value,
      state: fields.state.value,
      domain: fields.domain.value,
      mode: getSelectedMode(),
      startDate: fields.startDate.value,
      endDate: fields.endDate.value,
      projectTitle: fields.projectTitle.value
    };

    const payload = {
      step: currentStep,
      data: data
    };

    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  }

  function loadProgress() {
    const rawData = localStorage.getItem(STORAGE_KEY);
    if (!rawData) return;

    try {
      const payload = JSON.parse(rawData);
      const data = payload.data || {};

      Object.keys(fields).forEach(key => {
        if (data[key] && fields[key]) {
          fields[key].value = data[key];
          validateField(key);
        }
      });

      if (data.mode) {
        const radio = document.querySelector(`input[name="mode"][value="${data.mode}"]`);
        if (radio) radio.checked = true;
      }

      handleDateChange();
      goToStep(payload.step || 1);
    } catch (e) {
      console.warn('Error loading progress', e);
    }
  }

  function clearProgress() {
    localStorage.removeItem(STORAGE_KEY);
    form.reset();

    // Clear validation styling
    document.querySelectorAll('.input-wrapper').forEach(w => {
      w.classList.remove('is-valid', 'has-error');
    });

    document.querySelectorAll('.error-msg').forEach(e => {
      e.textContent = '';
    });

    durationVal.textContent = 'Select start and end dates';
    durationBox.style.borderColor = 'var(--accent-cyan)';

    // Reset view
    successScreen.classList.add('hidden');
    form.classList.remove('hidden');
    document.querySelector('.form-header-bar').classList.remove('hidden');
    document.querySelector('.progress-track').classList.remove('hidden');

    goToStep(1);
  }

  // ==========================================================================
  // 6. SUCCESS SCREEN, API SUBMISSION & CONFETTI
  // ==========================================================================
  async function handleFormSubmission() {
    const batchInput = document.getElementById('batchCode');
    const batchCode = batchInput ? batchInput.value.trim() : 'APP26-27';

    const payload = {
      batchCode: batchCode,
      fullName: fields.fullName.value.trim(),
      registerNo: fields.registerNo.value.trim(),
      collegeName: fields.collegeName.value.trim(),
      degreeBranch: fields.degreeBranch.value.trim(),
      state: fields.state.value,
      email: fields.email.value.trim(),
      phone: fields.phone.value.trim(),
      domain: fields.domain.value,
      mode: getSelectedMode(),
      startDate: fields.startDate.value,
      endDate: fields.endDate.value,
      projectTitle: fields.projectTitle.value.trim()
    };

    btnSubmit.disabled = true;
    const originalSubmitHtml = btnSubmit.innerHTML;
    btnSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Submitting...</span>';

    try {
      const response = await fetch('/api/students', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const resData = await response.json();
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = originalSubmitHtml;

      if (response.ok && resData.success) {
        showSuccess(resData.reference_id);
      } else {
        alert(resData.error || 'Submission could not be saved. Please check all fields.');
      }
    } catch (err) {
      console.warn('Backend submit fallback:', err);
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = originalSubmitHtml;
      showSuccess();
    }
  }

  function showSuccess(customRefId) {
    // Generate Reference ID or use returned ID from SQLite
    const refId = customRefId || generateReferenceId();
    successRefId.textContent = refId;

    // Populate Certificate Card
    certStudentName.textContent = fields.fullName.value;
    certDomainTag.textContent = fields.domain.value;
    certModeTag.textContent = getSelectedMode();
    certCollege.textContent = fields.collegeName.value;
    certRegno.textContent = fields.registerNo.value;

    const dur = calculateDuration();
    certDurationVal.textContent = dur ? dur.text : '-';

    const now = new Date();
    certDateStamp.textContent = `Submitted on ${now.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;

    // Hide form & header
    form.classList.add('hidden');
    document.querySelector('.form-header-bar').classList.add('hidden');
    document.querySelector('.progress-track').classList.add('hidden');

    // Show success screen
    successScreen.classList.remove('hidden');

    // Trigger Confetti Celebration
    createConfetti();
  }

  function generateReferenceId() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let code = '';
    for (let i = 0; i < 6; i++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return `AP-2026-${code}`;
  }

  function createConfetti() {
    const container = document.getElementById('confetti-container');
    container.innerHTML = '';

    const colors = ['#2563EB', '#06B6D4', '#10B981', '#3B82F6', '#F59E0B', '#8B5CF6'];
    const particleCount = 70;

    for (let i = 0; i < particleCount; i++) {
      const piece = document.createElement('div');
      piece.className = 'confetti-piece';
      piece.style.left = `${Math.random() * 100}%`;
      piece.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];

      const size = Math.random() * 8 + 6;
      piece.style.width = `${size}px`;
      piece.style.height = `${size * (Math.random() > 0.5 ? 1 : 2.5)}px`;

      const duration = Math.random() * 2 + 1.8;
      const delay = Math.random() * 0.4;
      piece.style.animationDuration = `${duration}s`;
      piece.style.animationDelay = `${delay}s`;

      container.appendChild(piece);
    }

    setTimeout(() => {
      container.innerHTML = '';
    }, 4000);
  }

  // ==========================================================================
  // 7. KEYBOARD NAVIGATION ACCESSIBILITY
  // ==========================================================================
  function setupKeyboardListeners() {
    document.addEventListener('keydown', (e) => {
      // Avoid triggering when inside modal or success screen
      if (!successScreen.classList.contains('hidden')) return;

      if (e.key === 'Enter') {
        // Shift + Enter -> Go Back
        if (e.shiftKey) {
          e.preventDefault();
          if (currentStep > 1) goToStep(currentStep - 1);
        } else {
          // Enter -> Go Next (or Submit if step 5)
          e.preventDefault();
          if (currentStep < totalSteps) {
            if (validateStep(currentStep)) {
              goToStep(currentStep + 1);
            }
          } else if (currentStep === totalSteps) {
            if (validateStep(currentStep)) {
              showSuccess();
            }
          }
        }
      }
    });
  }

  // Initialize Application
  initializeForm();
});
