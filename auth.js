/* Microsoft Entra External ID authentication for the Sustainability Quiz */
const AUTH_SETTINGS = {
  clientId: "be527a41-15ad-4dbd-b705-231535f83d39",
  tenantId: "ffd10b73-d357-4068-ab82-b941b9b0279e",
  ciamHost: "quizaccounts.ciamlogin.com",
  redirectUri: "https://tvd2804.github.io/Sustainability-quiz-TEST/"
};

const msalConfig = {
  auth: {
    clientId: AUTH_SETTINGS.clientId,
    authority: `https://${AUTH_SETTINGS.ciamHost}/${AUTH_SETTINGS.tenantId}`,
    knownAuthorities: [AUTH_SETTINGS.ciamHost],
    redirectUri: AUTH_SETTINGS.redirectUri,
    postLogoutRedirectUri: AUTH_SETTINGS.redirectUri,
    navigateToLoginRequestUrl: true
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const loginRequest = { scopes: ["openid", "profile", "email"] };
const QUIZ_PROFILE_PREFIX = "stockmeier-quiz-profile-";
let msalInstance = null;
let currentAccount = null;

function setAuthStatus(message, isError = false) {
  const status = document.getElementById("authStatus");
  if (!status) return;
  status.textContent = message || "";
  status.classList.toggle("error", Boolean(isError));
}

function getQuizUserProfile() {
  if (!currentAccount) return null;
  const claims = currentAccount.idTokenClaims || {};
  return {
    userId: claims.sub || claims.oid || currentAccount.homeAccountId || "",
    email: claims.email || (Array.isArray(claims.emails) ? claims.emails[0] : "") || currentAccount.username || "",
    displayName: claims.name || currentAccount.name || ""
  };
}

function profileStorageKey() {
  const profile = getQuizUserProfile();
  return profile?.userId ? QUIZ_PROFILE_PREFIX + profile.userId : "";
}

function getSavedQuizProfile() {
  const key = profileStorageKey();
  if (!key) return {};
  try { return JSON.parse(localStorage.getItem(key) || "{}"); }
  catch (error) { return {}; }
}

function saveQuizProfile(profile) {
  const key = profileStorageKey();
  if (!key) return;
  const safeProfile = {
    country: String(profile?.country || ""),
    location: String(profile?.location || ""),
    department: String(profile?.department || "")
  };
  localStorage.setItem(key, JSON.stringify(safeProfile));
}

function updateAuthenticationInterface() {
  const gate = document.getElementById("authGate");
  const app = document.getElementById("quizApplication");
  const user = document.getElementById("loggedInUser");
  const logout = document.getElementById("logoutButton");
  const login = document.getElementById("entraLoginButton");
  if (currentAccount) {
    gate?.classList.add("hidden");
    app?.classList.remove("hidden");
    logout?.classList.remove("hidden");
    if (login) login.disabled = false;
    const profile = getQuizUserProfile();
    if (user) user.textContent = profile?.displayName || profile?.email || "Signed in";
    setAuthStatus("");
  } else {
    gate?.classList.remove("hidden");
    app?.classList.add("hidden");
    logout?.classList.add("hidden");
    if (user) user.textContent = "";
    if (login) login.disabled = false;
    setAuthStatus("Log in or create an account to continue.");
  }
}

async function initialiseEntraAuthentication() {
  if (typeof msal === "undefined" || !msal.PublicClientApplication) {
    throw new Error("The Microsoft authentication library could not be loaded.");
  }
  setAuthStatus("Checking sign-in status...");
  msalInstance = new msal.PublicClientApplication(msalConfig);
  const redirectResponse = await msalInstance.handleRedirectPromise();
  if (redirectResponse?.account) {
    currentAccount = redirectResponse.account;
    msalInstance.setActiveAccount(currentAccount);
  } else {
    const accounts = msalInstance.getAllAccounts();
    currentAccount = accounts[0] || null;
    if (currentAccount) msalInstance.setActiveAccount(currentAccount);
  }
  updateAuthenticationInterface();
  return currentAccount;
}

async function signInToQuiz() {
  try {
    if (!msalInstance) await initialiseEntraAuthentication();
    const button = document.getElementById("entraLoginButton");
    if (button) button.disabled = true;
    setAuthStatus("Opening secure sign-in...");
    await msalInstance.loginRedirect(loginRequest);
  } catch (error) {
    console.error("Sign-in failed:", error);
    setAuthStatus(error.message || "Sign-in failed.", true);
    const button = document.getElementById("entraLoginButton");
    if (button) button.disabled = false;
  }
}

async function signOutOfQuiz() {
  if (!msalInstance) return;
  await msalInstance.logoutRedirect({
    account: msalInstance.getActiveAccount() || currentAccount,
    postLogoutRedirectUri: AUTH_SETTINGS.redirectUri
  });
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("entraLoginButton")?.addEventListener("click", signInToQuiz);
  document.getElementById("logoutButton")?.addEventListener("click", signOutOfQuiz);
});

window.quizAuthentication = {
  initialise: initialiseEntraAuthentication,
  signIn: signInToQuiz,
  signOut: signOutOfQuiz,
  getProfile: getQuizUserProfile,
  getSavedQuizProfile,
  saveQuizProfile,
  isAuthenticated: () => Boolean(currentAccount)
};
