# restore-state-v2.ps1 – Surgically inserts missing state at the correct location
$file = "src\App.tsx"
$content = [System.IO.File]::ReadAllText((Resolve-Path $file), [System.Text.Encoding]::UTF8)

# Exact string to find and expand:
$oldSection = "  const [shareModalOpen, setShareModalOpen] = useState(false);`r`n  const [logisticsAnalytics"
$newSection = @"
  const [shareModalOpen, setShareModalOpen] = useState(false);
  const [versionUploading, setVersionUploading] = useState(false);

  // Admin Portal State
  const [adminModalOpen, setAdminModalOpen] = useState(false);
  const [adminTab, setAdminTab] = useState<"users" | "payments" | "plans" | "analytics" | "settings" | "activity">("users");
  const [adminUsers, setAdminUsers] = useState<User[]>([]);
  const [adminPayments, setAdminPayments] = useState<any[]>([]);
  const [adminPlans, setAdminPlans] = useState<Plan[]>([]);
  const [adminLogs, setAdminLogs] = useState<AdminLog[]>([]);
  const [adminSearch, setAdminSearch] = useState("");
  const [adminLoading, setAdminLoading] = useState(false);
  const [adminPaymentFilter, setAdminPaymentFilter] = useState<"all" | "success" | "pending_verification" | "rejected" | "refunded">("all");
  const [adminUserFilter, setAdminUserFilter] = useState<"all" | "pro" | "free">("all");

  // Admin Edit User
  const [adminEditUser, setAdminEditUser] = useState<User | null>(null);
  const [adminEditUserName, setAdminEditUserName] = useState("");
  const [adminEditUserEmail, setAdminEditUserEmail] = useState("");
  const [adminEditUserMobile, setAdminEditUserMobile] = useState("");
  const [adminEditUserIsPro, setAdminEditUserIsPro] = useState(false);
  const [adminEditUserOpen, setAdminEditUserOpen] = useState(false);

  // Admin Plans
  const [adminEditPlan, setAdminEditPlan] = useState<Plan | null>(null);
  const [adminPlanName, setAdminPlanName] = useState("");
  const [adminPlanPrice, setAdminPlanPrice] = useState(0);
  const [adminPlanPeriod, setAdminPlanPeriod] = useState<Plan["billingPeriod"]>("monthly");
  const [adminPlanFeatureInput, setAdminPlanFeatureInput] = useState("");
  const [adminPlanFeatures, setAdminPlanFeatures] = useState<string[]>([]);
  const [adminPlanActive, setAdminPlanActive] = useState(true);
  const [adminPlanModalOpen, setAdminPlanModalOpen] = useState(false);
  // Extended plan editor fields
  const [adminPlanDescription, setAdminPlanDescription] = useState("");
  const [adminPlanHighlighted, setAdminPlanHighlighted] = useState(false);
  const [adminPlanColor, setAdminPlanColor] = useState("#f59e0b");
  const [adminPlanMaxReports, setAdminPlanMaxReports] = useState(0);
  const [adminPlanSortOrder, setAdminPlanSortOrder] = useState(99);

  // Checkout: plans loaded from DB + selected plan
  const [checkoutPlans, setCheckoutPlans] = useState<Plan[]>([]);
  const [selectedPlanId, setSelectedPlanId] = useState<string | null>(null);

  // Admin Settings -- feature flags
  const [adminFeatureAI, setAdminFeatureAI] = useState(true);
  const [adminFeatureUPI, setAdminFeatureUPI] = useState(true);
  const [adminFeatureGoogleLogin, setAdminFeatureGoogleLogin] = useState(true);
  const [adminMaintenanceMode, setAdminMaintenanceMode] = useState(false);
  const [globalFreeLimit, setGlobalFreeLimit] = useState(() => {
    if (typeof window === "undefined") return 3;
    const stored = window.localStorage.getItem("sheetcodecrest_global_free_limit");
    return stored ? Number(stored) : 3;
  });

  const [mockupTabActive, setMockupTabActive] = useState<"shopify" | "logistics" | "universal">("shopify");

  // Universal Profiler Data
  const [dataProfile, setDataProfile] = useState<DataProfile | null>(null);

  // Logistics Optimizer Data
  const [logisticsAnalytics
"@

if ($content.Contains($oldSection)) {
    $content = $content.Replace($oldSection, $newSection)
    [System.IO.File]::WriteAllText((Resolve-Path $file), $content, [System.Text.Encoding]::UTF8)
    Write-Host "SUCCESS: State declarations inserted."
} else {
    Write-Host "Pattern not found. Checking CRLF/LF differences..."
    # Try LF-only
    $oldSectionLF = "  const [shareModalOpen, setShareModalOpen] = useState(false);`n  const [logisticsAnalytics"
    if ($content.Contains($oldSectionLF)) {
        $newSectionLF = $newSection -replace "`r`n", "`n"
        $content = $content.Replace($oldSectionLF, $newSectionLF)
        [System.IO.File]::WriteAllText((Resolve-Path $file), $content, [System.Text.Encoding]::UTF8)
        Write-Host "SUCCESS (LF): State declarations inserted."
    } else {
        # Show what's around the marker
        $idx = $content.IndexOf("setShareModalOpen")
        if ($idx -ge 0) {
            Write-Host "Context around shareModalOpen:"
            Write-Host $content.Substring($idx, [Math]::Min(200, $content.Length - $idx)) | Format-Hex
        } else {
            Write-Host "shareModalOpen not found at all!"
        }
    }
}
