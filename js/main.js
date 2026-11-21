/**
 * Maribel Bravo - Personal Portfolio
 * Main JavaScript (js/main.js)
 */

document.addEventListener("DOMContentLoaded", () => {
  // ── 1. TYPEWRITER EFFECT ──────────────────────────────────
  const words = ["Aspiring Developer", "Problem Solver", "Tech Enthusiast", "Lifelong Learner"];
  let wordIndex = 0;
  let charIndex = 0;
  let isDeleting = false;
  const typedOutput = document.getElementById("typed-output");

  function typeEffect() {
    if (!typedOutput) return;

    const currentWord = words[wordIndex];
    if (isDeleting) {
      typedOutput.textContent = currentWord.substring(0, charIndex - 1);
      charIndex--;
    } else {
      typedOutput.textContent = currentWord.substring(0, charIndex + 1);
      charIndex++;
    }

    let speed = isDeleting ? 40 : 80;

    if (!isDeleting && charIndex === currentWord.length) {
      speed = 1800; // Pause at full word
      isDeleting = true;
    } else if (isDeleting && charIndex === 0) {
      isDeleting = false;
      wordIndex = (wordIndex + 1) % words.length;
      speed = 400;
    }

    setTimeout(typeEffect, speed);
  }
  typeEffect();

  // ── 2. NAVBAR SCROLL & READING PROGRESS & BACK TO TOP ─────
  const navbar = document.querySelector(".custom-navbar");
  const progressBar = document.getElementById("scroll-progress");
  const backToTopBtn = document.getElementById("back-to-top");

  window.addEventListener("scroll", () => {
    const scrollY = window.scrollY;
    const docHeight = document.documentElement.scrollHeight - window.innerHeight;
    const progress = docHeight > 0 ? (scrollY / docHeight) * 100 : 0;

    if (progressBar) {
      progressBar.style.width = `${progress}%`;
    }

    if (navbar) {
      if (scrollY > 50) {
        navbar.classList.add("scrolled");
      } else {
        navbar.classList.remove("scrolled");
      }
    }

    if (backToTopBtn) {
      if (scrollY > 400) {
        backToTopBtn.classList.add("show");
      } else {
        backToTopBtn.classList.remove("show");
      }
    }
  });

  // ── 3. SCROLL REVEAL OBSERVER ─────────────────────────────
  const reveals = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window) {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("active");
          }
        });
      },
      { threshold: 0.12 }
    );

    reveals.forEach((el) => observer.observe(el));
  } else {
    // Fallback if IntersectionObserver is not supported
    reveals.forEach((el) => el.classList.add("active"));
  }

  // ── 4. NAVBAR SCROLLSPY (ACTIVE LINK) ─────────────────────
  const spySections = document.querySelectorAll("section[id]");
  const spyLinks = document.querySelectorAll(".navbar-nav .nav-link");

  function setActiveLink(id) {
    spyLinks.forEach((link) => {
      link.classList.toggle("active", link.getAttribute("href") === `#${id}`);
    });
  }

  if ("IntersectionObserver" in window && spySections.length) {
    const spyObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            setActiveLink(entry.target.id);
          }
        });
      },
      { rootMargin: "-35% 0px -55% 0px" }
    );
    spySections.forEach((section) => spyObserver.observe(section));
  }

  spyLinks.forEach((link) => {
    link.addEventListener("click", () => setActiveLink(link.getAttribute("href").slice(1)));
  });

  // ── 5. CLOSE MOBILE MENU ON LINK CLICK ────────────────────
  const navLinks = document.querySelectorAll(".navbar-nav .nav-link");
  const navCollapse = document.getElementById("navbarContent");
  if (navCollapse && typeof bootstrap !== "undefined" && bootstrap.Collapse) {
    navLinks.forEach((link) => {
      link.addEventListener("click", () => {
        if (navCollapse.classList.contains("show")) {
          const bsCollapse = bootstrap.Collapse.getInstance(navCollapse) || new bootstrap.Collapse(navCollapse);
          bsCollapse.hide();
        }
      });
    });
  }

  // ── 7. FORMSPREE CONTACT FORM ─────────────────────────────
  const contactForm = document.getElementById("contact-form");
  if (contactForm) {
    contactForm.addEventListener("submit", async function (e) {
      e.preventDefault();

      const btn = document.querySelector(".btn-send");
      const msg = document.getElementById("contact-msg");
      const originalHtml = btn ? btn.innerHTML : "Send";

      if (btn) {
        btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Sending...';
        btn.disabled = true;
      }

      try {
        const res = await fetch("https://formspree.io/f/xojbkprk", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
          },
          body: JSON.stringify({
            name: document.querySelector('[name="name"]').value,
            email: document.querySelector('[name="email"]').value,
            message: document.querySelector('[name="message"]').value,
          }),
        });

        if (res.ok) {
          if (msg) {
            msg.textContent = "✅ Message sent successfully! I'll get back to you soon.";
            msg.className = "contact-msg success";
          }
          contactForm.reset();
        } else {
          if (msg) {
            msg.textContent = "❌ Something went wrong. Please try again or email directly.";
            msg.className = "contact-msg error";
          }
        }
      } catch {
        if (msg) {
          msg.textContent = "❌ Failed to connect. Please check your network and try again.";
          msg.className = "contact-msg error";
        }
      } finally {
        if (btn) {
          btn.innerHTML = originalHtml;
          btn.disabled = false;
        }
      }
    });
  }
});
