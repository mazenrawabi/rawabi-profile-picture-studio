import { useCallback, useEffect, useRef, useState } from "react";
import { authDisabled, isAuthConfigured, loginRequest, msal } from "./auth";

const WIDTH = 605;
const HEIGHT = 688;
const GUIDE_Y = 63;
const ARTWORK_SOURCE_SIZE = 950;

function imageFromBlob(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Image could not be opened"));
    };
    image.src = url;
  });
}

function AuthGate({ children }) {
  const [status, setStatus] = useState(authDisabled ? "authenticated" : "loading");
  const [account, setAccount] = useState(null);
  const [error, setError] = useState("");

  useEffect(() => {
    if (authDisabled || !isAuthConfigured || !msal) {
      setStatus(authDisabled ? "authenticated" : "configuration");
      return;
    }

    let active = true;
    (async () => {
      try {
        await msal.initialize();
        const redirect = await msal.handleRedirectPromise();
        const signedIn = redirect?.account ?? msal.getAllAccounts()[0] ?? null;
        if (signedIn) msal.setActiveAccount(signedIn);
        if (active) {
          setAccount(signedIn);
          setStatus(signedIn ? "authenticated" : "anonymous");
        }
      } catch (authError) {
        if (active) {
          setError(authError instanceof Error ? authError.message : String(authError));
          setStatus("anonymous");
        }
      }
    })();
    return () => { active = false; };
  }, []);

  async function signIn() {
    setError("");
    try {
      await msal.loginRedirect(loginRequest);
    } catch (authError) {
      setError(authError instanceof Error ? authError.message : String(authError));
    }
  }

  async function signOut() {
    await msal.logoutRedirect({ account });
  }

  if (status === "loading") {
    return <AuthScreen title="Checking your Rawabi account…" busy />;
  }
  if (status === "configuration") {
    return (
      <AuthScreen
        title="Entra configuration required"
        description="The application is ready, but its Microsoft Entra application ID has not been added yet."
        error="Add VITE_ENTRA_CLIENT_ID to the deployment environment."
      />
    );
  }
  if (status === "anonymous") {
    return (
      <AuthScreen
        title="Employee sign in"
        description="Use your Rawabi Microsoft account to access the Profile Picture Studio."
        action="Sign in with Microsoft"
        onAction={signIn}
        error={error}
      />
    );
  }

  return children({
    account,
    signOut: authDisabled ? null : signOut,
  });
}

function AuthScreen({ title, description, action, onAction, busy, error }) {
  return (
    <main className="auth-shell">
      <section className="auth-card">
        <img className="auth-logo" src="/rawabi-logo.svg" alt="Rawabi Holding" />
        <p className="eyebrow">Human Capital</p>
        <h1>{title}</h1>
        {description && <p className="auth-description">{description}</p>}
        {action && <button className="auth-button" onClick={onAction}>{action}</button>}
        {busy && <div className="auth-security"><span />Secure sign-in</div>}
        {error && <p className="auth-error">{error}</p>}
      </section>
    </main>
  );
}

function Studio({ account, signOut }) {
  const canvasRef = useRef(null);
  const fileRef = useRef(null);
  const dragRef = useRef(null);
  const brushRef = useRef(null);
  const originalRef = useRef(null);
  const maskRef = useRef(null);
  const originalMaskRef = useRef(null);
  const compositeRef = useRef(null);

  const [photo, setPhoto] = useState(null);
  const [, refreshComposite] = useState(0);
  const [artwork, setArtwork] = useState(null);
  const [filename, setFilename] = useState("");
  const [sourceFile, setSourceFile] = useState(null);
  const [processing, setProcessing] = useState("idle");
  const [progress, setProgress] = useState(0);
  const [processingError, setProcessingError] = useState("");
  const [tool, setTool] = useState("photo");
  const [photoX, setPhotoX] = useState(WIDTH / 2);
  const [photoY, setPhotoY] = useState(HEIGHT / 2);
  const [photoScale, setPhotoScale] = useState(1);
  const [artworkX, setArtworkX] = useState(-290);
  const [artworkY, setArtworkY] = useState(322);
  const [artworkScale, setArtworkScale] = useState(0.61);
  const [brushSize, setBrushSize] = useState(34);
  const [dragging, setDragging] = useState(false);

  const baseScale = photo ? Math.max(WIDTH / photo.width, HEIGHT / photo.height) : 1;

  useEffect(() => {
    const overlay = new Image();
    overlay.onload = () => setArtwork(overlay);
    overlay.src = "/brand-overlay.png";
  }, []);

  const drawCanvas = useCallback((canvas, exportMode = false) => {
    const context = canvas.getContext("2d");
    if (!context) return;
    context.clearRect(0, 0, WIDTH, HEIGHT);
    context.fillStyle = "#8d939e";
    context.fillRect(0, 0, WIDTH, HEIGHT);

    if (photo) {
      const scale = baseScale * photoScale;
      const width = photo.width * scale;
      const height = photo.height * scale;
      context.drawImage(photo, photoX - width / 2, photoY - height / 2, width, height);
    }

    context.strokeStyle = "#858b96";
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(0, GUIDE_Y + 0.5);
    context.lineTo(WIDTH, GUIDE_Y + 0.5);
    context.stroke();

    if (artwork) {
      const size = ARTWORK_SOURCE_SIZE * artworkScale;
      context.drawImage(artwork, artworkX, artworkY, size, size);
    }

    if (!exportMode && photo) {
      context.save();
      context.strokeStyle = tool === "artwork" ? "rgba(255,255,255,.82)" : "rgba(31,79,66,.72)";
      context.lineWidth = 1.5;
      context.setLineDash([6, 5]);
      if (tool === "artwork") {
        context.strokeRect(artworkX, artworkY, ARTWORK_SOURCE_SIZE * artworkScale, ARTWORK_SOURCE_SIZE * artworkScale);
      } else {
        const scale = baseScale * photoScale;
        context.strokeRect(
          photoX - (photo.width * scale) / 2,
          photoY - (photo.height * scale) / 2,
          photo.width * scale,
          photo.height * scale,
        );
      }
      context.restore();
    }
  }, [artwork, artworkScale, artworkX, artworkY, baseScale, photo, photoScale, photoX, photoY, tool]);

  useEffect(() => {
    if (canvasRef.current) drawCanvas(canvasRef.current);
  }, [drawCanvas]);

  function resetPlacement() {
    setPhotoX(WIDTH / 2);
    setPhotoY(HEIGHT / 2);
    setPhotoScale(1);
    setArtworkX(-290);
    setArtworkY(322);
    setArtworkScale(0.61);
  }

  function rebuildComposite() {
    const original = originalRef.current;
    const mask = maskRef.current;
    const composite = compositeRef.current;
    if (!original || !mask || !composite) return;
    const context = composite.getContext("2d");
    if (!context) return;
    context.clearRect(0, 0, composite.width, composite.height);
    context.globalCompositeOperation = "source-over";
    context.drawImage(original, 0, 0, composite.width, composite.height);
    context.globalCompositeOperation = "destination-in";
    context.drawImage(mask, 0, 0);
    context.globalCompositeOperation = "source-over";
    refreshComposite((value) => value + 1);
  }

  function resetCutout() {
    const mask = maskRef.current;
    const originalMask = originalMaskRef.current;
    if (!mask || !originalMask) return;
    const context = mask.getContext("2d");
    if (!context) return;
    context.clearRect(0, 0, mask.width, mask.height);
    context.drawImage(originalMask, 0, 0);
    rebuildComposite();
  }

  function brushAt(x, y) {
    if (!photo || (tool !== "restore" && tool !== "erase")) return;
    const scale = baseScale * photoScale;
    const sourceX = (x - (photoX - (photo.width * scale) / 2)) / scale;
    const sourceY = (y - (photoY - (photo.height * scale) / 2)) / scale;
    if (sourceX < 0 || sourceY < 0 || sourceX > photo.width || sourceY > photo.height) return;

    const radius = brushSize / scale;
    const mask = maskRef.current;
    const context = mask?.getContext("2d");
    if (!mask || !context) return;
    const gradient = context.createRadialGradient(sourceX, sourceY, radius * 0.55, sourceX, sourceY, radius);
    gradient.addColorStop(0, "rgba(255,255,255,1)");
    gradient.addColorStop(1, "rgba(255,255,255,0)");
    context.save();
    context.globalCompositeOperation = tool === "restore" ? "source-over" : "destination-out";
    context.fillStyle = gradient;
    context.beginPath();
    context.arc(sourceX, sourceY, radius, 0, Math.PI * 2);
    context.fill();
    context.restore();
    rebuildComposite();
  }

  async function processFile(file) {
    if (!file || !file.type.startsWith("image/")) return;
    setSourceFile(file);
    setFilename(file.name);
    setPhoto(null);
    setProcessing("processing");
    setProgress(2);
    setProcessingError("");

    try {
      const { removeBackground } = await import("@imgly/background-removal");
      const cutoutBlob = await removeBackground(file, {
        device: "cpu",
        model: "isnet",
        output: { format: "image/png", quality: 1 },
        progress: (_key, current, total) => {
          if (total > 0) setProgress(Math.max(3, Math.min(92, Math.round((current / total) * 92))));
        },
      });
      setProgress(96);
      const [original, cutout] = await Promise.all([imageFromBlob(file), imageFromBlob(cutoutBlob)]);
      const width = cutout.naturalWidth;
      const height = cutout.naturalHeight;
      const mask = document.createElement("canvas");
      const originalMask = document.createElement("canvas");
      const composite = document.createElement("canvas");
      mask.width = originalMask.width = composite.width = width;
      mask.height = originalMask.height = composite.height = height;
      mask.getContext("2d")?.drawImage(cutout, 0, 0, width, height);
      originalMask.getContext("2d")?.drawImage(mask, 0, 0);
      originalRef.current = original;
      maskRef.current = mask;
      originalMaskRef.current = originalMask;
      compositeRef.current = composite;
      const context = composite.getContext("2d");
      if (!context) throw new Error("Canvas is not available");
      context.drawImage(original, 0, 0, width, height);
      context.globalCompositeOperation = "destination-in";
      context.drawImage(mask, 0, 0);
      context.globalCompositeOperation = "source-over";
      setPhoto(composite);
      resetPlacement();
      setTool("photo");
      setProgress(100);
      setProcessing("ready");
    } catch (processError) {
      console.error("Background removal failed:", processError);
      setProcessing("error");
      setProcessingError("Background removal didn’t finish. Please try again; if it continues, use a smaller JPG or PNG.");
    }
  }

  function canvasPoint(event) {
    const bounds = event.currentTarget.getBoundingClientRect();
    return {
      x: ((event.clientX - bounds.left) * WIDTH) / bounds.width,
      y: ((event.clientY - bounds.top) * HEIGHT) / bounds.height,
    };
  }

  function pointerDown(event) {
    if (!photo) return;
    const point = canvasPoint(event);
    if (tool === "restore" || tool === "erase") {
      brushRef.current = point;
      event.currentTarget.setPointerCapture(event.pointerId);
      brushAt(point.x, point.y);
      return;
    }
    dragRef.current = {
      x: point.x,
      y: point.y,
      startX: tool === "photo" ? photoX : artworkX,
      startY: tool === "photo" ? photoY : artworkY,
    };
    event.currentTarget.setPointerCapture(event.pointerId);
    setDragging(true);
  }

  function pointerMove(event) {
    if (brushRef.current && (tool === "restore" || tool === "erase")) {
      const point = canvasPoint(event);
      const previous = brushRef.current;
      const distance = Math.hypot(point.x - previous.x, point.y - previous.y);
      const steps = Math.max(1, Math.ceil(distance / Math.max(4, brushSize * 0.25)));
      for (let index = 1; index <= steps; index += 1) {
        brushAt(previous.x + ((point.x - previous.x) * index) / steps, previous.y + ((point.y - previous.y) * index) / steps);
      }
      brushRef.current = point;
      return;
    }
    if (!dragRef.current) return;
    const point = canvasPoint(event);
    const dx = point.x - dragRef.current.x;
    const dy = point.y - dragRef.current.y;
    if (tool === "photo") {
      setPhotoX(dragRef.current.startX + dx);
      setPhotoY(dragRef.current.startY + dy);
    } else {
      setArtworkX(dragRef.current.startX + dx);
      setArtworkY(dragRef.current.startY + dy);
    }
  }

  function pointerUp() {
    dragRef.current = null;
    brushRef.current = null;
    setDragging(false);
  }

  function nudge(dx, dy) {
    if (tool === "photo") {
      setPhotoX((value) => value + dx);
      setPhotoY((value) => value + dy);
    } else {
      setArtworkX((value) => value + dx);
      setArtworkY((value) => value + dy);
    }
  }

  function download() {
    if (!photo) return;
    const output = document.createElement("canvas");
    output.width = WIDTH;
    output.height = HEIGHT;
    drawCanvas(output, true);
    output.toBlob((blob) => {
      if (!blob) return;
      const link = document.createElement("a");
      const safeName = filename.replace(/\.[^.]+$/, "").replace(/[^a-z0-9-_@.]+/gi, "-") || "employee";
      link.href = URL.createObjectURL(blob);
      link.download = `${safeName}-profile.png`;
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 500);
    }, "image/png");
  }

  const accountLabel = account?.name || account?.username;

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="rawabi-lockup"><img src="/rawabi-logo.svg" alt="Rawabi Holding" /></div>
        <div className="product-lockup"><p className="eyebrow">Human Capital</p><h1>Profile Picture Studio</h1></div>
        {accountLabel ? (
          <div className="account-menu">
            <div className="account-copy"><strong>{account?.name || "Rawabi employee"}</strong><span>{account?.username}</span></div>
            {signOut && <button className="signout-button" onClick={signOut}>Sign out</button>}
          </div>
        ) : <div className="privacy-pill"><span />Processed on this device</div>}
      </header>

      <section className="workspace portrait-workspace">
        <aside className="control-panel">
          <div className="panel-section upload-section">
            <div className="step-number">01</div><h2>Employee photo</h2>
            <p>Choose a clear portrait. The background will be removed automatically on this device.</p>
            <input ref={fileRef} className="sr-only" type="file" accept="image/png,image/jpeg,image/webp" onChange={(event) => processFile(event.target.files?.[0])} />
            <button className="upload-button" disabled={processing === "processing"} onClick={() => fileRef.current?.click()}>
              <span>＋</span>{processing === "processing" ? "Removing background…" : photo ? "Replace photo" : "Choose photo"}
            </button>
            {processing === "processing" && <div className="processing-status" role="status" aria-live="polite"><div className="progress-track"><span style={{ width: `${progress}%` }} /></div><p>Preparing cutout · {progress}%</p><small>The first photo may take a little longer while the removal model loads.</small></div>}
            {processing === "ready" && <p className="file-name" title={filename}>✓ High-quality cutout ready · {filename}</p>}
            {processing === "error" && <div className="processing-error" role="alert"><p>{processingError}</p><button onClick={() => processFile(sourceFile)}>Try again</button></div>}
          </div>

          <div className={`panel-section adjustments ${photo ? "" : "is-disabled"}`}>
            <div className="step-number">02</div>
            <div className="section-title"><h2>Fine tune</h2><button onClick={resetPlacement} disabled={!photo}>Reset</button></div>
            <label className="control-label">Edit</label>
            <div className="segmented edit-tools" role="group" aria-label="Choose editing tool">
              {[["photo", "Photo"], ["artwork", "Circle artwork"], ["restore", "Restore"], ["erase", "Erase"]].map(([value, label]) => <button key={value} className={tool === value ? "active" : ""} onClick={() => setTool(value)}>{label}</button>)}
            </div>
            {tool === "restore" || tool === "erase" ? <>
              <div className="cleanup-note"><strong>White clothing missing?</strong> Brush over it with Restore. Use Erase if any background returns.</div>
              <label className="control-label" htmlFor="brush-size">Brush size <span>{brushSize}px</span></label>
              <input id="brush-size" type="range" min="12" max="72" step="2" value={brushSize} onChange={(event) => setBrushSize(Number(event.target.value))} />
              <button className="reset-cutout" onClick={resetCutout}>Reset cutout</button>
            </> : <>
              <label className="control-label" htmlFor="scale">{tool === "photo" ? "Photo size" : "Artwork size"}<span>{Math.round((tool === "photo" ? photoScale : artworkScale / 0.61) * 100)}%</span></label>
              {tool === "photo" ? <input id="scale" type="range" min="0.7" max="2.5" step="0.01" value={photoScale} onChange={(event) => setPhotoScale(Number(event.target.value))} /> : <input id="scale" type="range" min="0.5" max="0.75" step="0.005" value={artworkScale} onChange={(event) => setArtworkScale(Number(event.target.value))} />}
              <div className="nudge-grid" aria-label={`Nudge ${tool}`}><button className="up" onClick={() => nudge(0, -2)} aria-label="Move up">↑</button><button className="left" onClick={() => nudge(-2, 0)} aria-label="Move left">←</button><span>{tool === "photo" ? "Photo" : "Artwork"}</span><button className="right" onClick={() => nudge(2, 0)} aria-label="Move right">→</button><button className="down" onClick={() => nudge(0, 2)} aria-label="Move down">↓</button></div>
            </>}
          </div>
          <div className="format-note"><span>i</span><p>Background removal runs locally in this browser—employee photos are never uploaded. The grey background, guide bar and turquoise artwork are then added automatically.</p></div>
        </aside>

        <section className="editor-card portrait-card" aria-label="Profile picture editor">
          <div className="editor-heading"><div><p className="eyebrow">Live preview</p><h2>{processing === "processing" ? "Removing the background" : photo ? tool === "restore" ? "Brush missing clothing back in" : tool === "erase" ? "Brush remaining background away" : `Drag the ${tool === "photo" ? "photo" : "circle artwork"}` : "Position the portrait"}</h2></div><span className="dimensions">605 × 688 px</span></div>
          <div className={`canvas-wrap portrait-canvas ${dragging ? "dragging" : ""} ${tool === "restore" || tool === "erase" ? "brushing" : ""}`} onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); processFile(event.dataTransfer.files?.[0]); }}>
            <canvas ref={canvasRef} width={WIDTH} height={HEIGHT} tabIndex={photo ? 0 : -1} aria-label="Employee profile picture preview" onPointerDown={pointerDown} onPointerMove={pointerMove} onPointerUp={pointerUp} onPointerCancel={pointerUp} onKeyDown={(event) => { if (["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(event.key)) { event.preventDefault(); nudge(event.key === "ArrowLeft" ? -2 : event.key === "ArrowRight" ? 2 : 0, event.key === "ArrowUp" ? -2 : event.key === "ArrowDown" ? 2 : 0); } }} />
            <div className="guide-label portrait-guide">HEAD BELOW THIS LINE</div>
            {processing === "processing" ? <div className="canvas-processing" role="status"><span className="spinner" /><strong>Creating employee cutout</strong><small>{progress}% complete</small></div> : !photo && <button className="empty-prompt" onClick={() => fileRef.current?.click()}><strong>Drop a photo here</strong><span>or click to browse</span></button>}
          </div>
          <div className="editor-footer"><div className="placement-guide"><span>!</span><p><strong>Check the headroom</strong> The top of the employee’s head must stay below the thin guide bar.</p></div><button className="download-button" disabled={!photo || !artwork} onClick={download}>Download PNG <span>↓</span></button></div>
        </section>
      </section>
    </main>
  );
}

export default function App() {
  return <AuthGate>{(auth) => <Studio {...auth} />}</AuthGate>;
}
