import React, { useCallback, useMemo, useRef, useState, useEffect } from "react";
import { Upload, Loader2, AlertCircle, CheckCircle2, Eye, X, Camera, MapPin } from "lucide-react";
import type { FollowUpQuestion } from "../../types";
import { apiClient } from "../../api/client";

interface FileInputProps {
  question: FollowUpQuestion;
  value: any;
  onChange: (value: any) => void;
  readOnly?: boolean;
  isApplied?: boolean;
}

const FILE_ACCEPT_MAP: Record<string, string[]> = {
  image: ["image/*"],
  pdf: ["application/pdf", ".pdf"],
  excel: [
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls",
    ".xlsx",
  ],
  stp: [".stp", ".step"],
  pvz: [".pvz"],
  doc: [".doc", ".docx", "application/msword", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
};

const FILE_VALIDATORS: Record<string, (file: File) => boolean> = {
  image: (file) => file.type.startsWith("image/"),
  pdf: (file) =>
    file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"),
  excel: (file) => {
    const allowedMimes = new Set(
      FILE_ACCEPT_MAP.excel.filter((value) => value.startsWith("application/"))
    );
    const lowerName = file.name.toLowerCase();
    return (
      allowedMimes.has(file.type) ||
      lowerName.endsWith(".xls") ||
      lowerName.endsWith(".xlsx")
    );
  },
  stp: (file) =>
    file.name.toLowerCase().endsWith(".stp") ||
    file.name.toLowerCase().endsWith(".step"),
  pvz: (file) => file.name.toLowerCase().endsWith(".pvz"),
  doc: (file) => {
    const lowerName = file.name.toLowerCase();
    return (
      lowerName.endsWith(".doc") ||
      lowerName.endsWith(".docx") ||
      file.type === "application/msword" ||
      file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
  },
};

const FILE_TYPE_LABELS: Record<string, string> = {
  image: "image",
  pdf: "PDF",
  excel: "Excel",
  stp: "STP",
  pvz: "PVZ",
  doc: "Word",
};

const resolveFileName = (fileUrl: any) => {
  if (!fileUrl) {
    return "";
  }
  if (typeof fileUrl === 'object' && fileUrl.url) {
    fileUrl = fileUrl.url;
  }
  if (typeof fileUrl !== 'string') {
    return "Uploaded file";
  }
  if (fileUrl.startsWith("data:")) {
    return "Embedded file";
  }
  try {
    const url = new URL(fileUrl, "http://dummy.local");
    const segments = url.pathname.split("/").filter(Boolean);
    const fileName = segments[segments.length - 1] || "Uploaded file";
    return decodeURIComponent(fileName);
  } catch (error) {
    const parts = fileUrl.split("/");
    return decodeURIComponent(parts[parts.length - 1] || "Uploaded file");
  }
};

const isImageUrl = (fileUrl: any) => {
  if (!fileUrl) {
    return false;
  }
  if (typeof fileUrl === 'object' && fileUrl.url) {
    return isImageUrl(fileUrl.url);
  }
  if (typeof fileUrl !== 'string') {
    return false;
  }
  try {
    const parsed = JSON.parse(fileUrl);
    if (parsed && parsed.url) {
      return isImageUrl(parsed.url);
    }
  } catch {
    // Not JSON, continue with regular check
  }
  if (fileUrl.startsWith("data:")) {
    return fileUrl.startsWith("data:image");
  }
  return /\.(png|jpg|jpeg|gif|bmp|webp|svg)$/i.test(fileUrl);
};

const parseFileValue = (fileValue: any) => {
  if (typeof fileValue === 'object' && fileValue !== null) {
    if (fileValue.url && fileValue.location) {
      return {
        type: "camera",
        url: fileValue.url,
        location: fileValue.location,
        timestamp: fileValue.timestamp,
      };
    }
    if (fileValue.url) {
      return { type: "file", url: fileValue.url };
    }
  }

  if (typeof fileValue === 'string') {
    try {
      const parsed = JSON.parse(fileValue);
      if (parsed && parsed.url && parsed.location) {
        return {
          type: "camera",
          url: parsed.url,
          location: parsed.location,
          timestamp: parsed.timestamp,
        };
      }
    } catch {
      // Not JSON, regular file URL
    }
  }
  
  return {
    type: "file",
    url: typeof fileValue === 'string' ? fileValue : "",
  };
};

export default function FileInput({
  question,
  value,
  onChange,
  readOnly = false,
  isApplied = false,
}: FileInputProps) {
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [uploadProgress, setUploadProgress] = useState<{
    percentage: number;
    loaded: number;
    total: number;
    timeRemaining?: number;
    speed?: number;
  } | null>(null);
  const [uploadMode, setUploadMode] = useState<"file" | "camera" | null>(null);
  const [cameraStream, setCameraStream] = useState<MediaStream | null>(null);
  const [videoReady, setVideoReady] = useState(false);
  const [location, setLocation] = useState<{ lat: number; lng: number; accuracy: number } | null>(null);
  const [locationError, setLocationError] = useState<string | null>(null);
  const [locationLoading, setLocationLoading] = useState(false);
  const [showCapturePreview, setShowCapturePreview] = useState(false);
  const [capturedImageBlob, setCapturedImageBlob] = useState<Blob | null>(null);
  const [capturedImageUrl, setCapturedImageUrl] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  const fileName = useMemo(() => resolveFileName(value), [value]);
  const showImagePreview = useMemo(() => isImageUrl(value), [value]);

  const accept = useMemo(() => {
    if (!question.allowedFileTypes || question.allowedFileTypes.length === 0) {
      return undefined;
    }
    const values = question.allowedFileTypes
      .flatMap((type) => FILE_ACCEPT_MAP[type] ?? [])
      .filter(Boolean);
    if (values.length === 0) {
      return undefined;
    }
    return Array.from(new Set(values)).join(",");
  }, [question.allowedFileTypes]);

  const getLocation = useCallback(async () => {
    setLocationLoading(true);
    setLocationError(null);

    if (!navigator.geolocation) {
      setLocationError("Geolocation is not supported by your browser");
      setLocationLoading(false);
      return;
    }

    navigator.geolocation.getCurrentPosition(
      (position) => {
        setLocation({
          lat: position.coords.latitude,
          lng: position.coords.longitude,
          accuracy: Math.min(position.coords.accuracy, 1000),
        });
        setLocationLoading(false);
      },
      (error) => {
        let errorMsg = "Unable to get location";
        if (error.code === 1) {
          errorMsg = "Location permission denied";
        } else if (error.code === 2) {
          errorMsg = "Location unavailable";
        } else if (error.code === 3) {
          errorMsg = "Location request timed out";
        }
        setLocationError(errorMsg);
        setLocationLoading(false);
      },
      { timeout: 10000, enableHighAccuracy: true }
    );
  }, []);

  const startCamera = useCallback(async () => {
    try {
      setError(null);

      // Try with more flexible constraints
      const constraints = {
        video: {
          facingMode: { ideal: "environment" }, // Use ideal instead of exact
          width: { ideal: 1280 },
          height: { ideal: 720 }
        },
        audio: false
      };

      const stream = await navigator.mediaDevices.getUserMedia(constraints);

      setCameraStream(stream);

      // Wait for next render cycle to ensure video element exists
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          // Explicitly play the video
          videoRef.current.play().catch(e => {
            console.error("Video play error:", e);
            setError("Failed to play video: " + e.message);
          });
        }
      }, 100);

      await getLocation();
    } catch (err) {
      console.error("Camera error:", err);
      let errorMsg = "Failed to access camera";
      if (err.name === "NotAllowedError") {
        errorMsg = "Camera permission denied. Please allow camera access in browser settings.";
      } else if (err.name === "NotFoundError") {
        errorMsg = "No camera found on this device.";
      } else if (err.name === "NotReadableError") {
        errorMsg = "Camera is already in use by another application.";
      }
      setError(errorMsg);
    }
  }, [getLocation]);

  const stopCamera = useCallback(() => {
    if (cameraStream) {
      cameraStream.getTracks().forEach((track) => track.stop());
      setCameraStream(null);
      setVideoReady(false);
    }
  }, [cameraStream]);

  const capturePhoto = useCallback(() => {
    if (!videoRef.current || !canvasRef.current) {
      setError("Camera not ready");
      return;
    }

    try {
      setError(null);

      const video = videoRef.current;
      if (!video.srcObject || video.videoWidth === 0 || video.videoHeight === 0) {
        setError("Video stream not ready. Please wait a moment and try again.");
        return;
      }

      if (video.readyState !== video.HAVE_ENOUGH_DATA) {
        setError("Video not ready. Please wait a moment and try again.");
        return;
      }

      const context = canvasRef.current.getContext("2d");
      if (!context) {
        throw new Error("Failed to get canvas context");
      }

      canvasRef.current.width = videoRef.current.videoWidth;
      canvasRef.current.height = videoRef.current.videoHeight;
      context.drawImage(videoRef.current, 0, 0);

      canvasRef.current.toBlob((blob) => {
        if (!blob) {
          setError("Failed to capture image");
          return;
        }

        setCapturedImageBlob(blob);
        setCapturedImageUrl(URL.createObjectURL(blob));
        setShowCapturePreview(true);
      }, "image/jpeg", 0.9);
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "Failed to capture image";
      setError(errorMsg);
    }
  }, []);

  const uploadCapturedImage = useCallback(async () => {
    if (!capturedImageBlob || !location) {
      setError("Image or location not ready");
      return;
    }

    try {
      setUploading(true);
      setError(null);

      const file = new File([capturedImageBlob], `camera-${Date.now()}.jpg`, {
        type: "image/jpeg",
      });

      const result = await apiClient.uploadFile(
        file,
        "form",
        question.id,
        (progress) => {
          setUploadProgress(progress);
        }
      );

      const uploadedUrl = apiClient.resolveUploadedFileUrl(result);

      if (!uploadedUrl) {
        throw new Error("File upload did not return a valid URL");
      }

      const metadata = {
        url: uploadedUrl,
        location: {
          latitude: location.lat,
          longitude: location.lng,
          accuracy: location.accuracy,
        },
        timestamp: new Date().toISOString(),
      };


      const newFileValue = JSON.stringify(metadata);
      let existingFiles: any[] = [];
      if (Array.isArray(value)) {
        existingFiles = value;
      } else if (value) {
        try {
          const parsed = JSON.parse(value);
          if (Array.isArray(parsed)) {
            existingFiles = parsed;
          } else {
            existingFiles = [value];
          }
        } catch {
          existingFiles = [value];
        }
      }
      onChange([...existingFiles, newFileValue]);

      stopCamera();
      setShowCapturePreview(false);
      setCapturedImageBlob(null);
      setCapturedImageUrl(null);
      setUploadMode(null);
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "Failed to upload image";
      setError(errorMsg);
    } finally {
      setUploading(false);
      setUploadProgress(null);
    }
  }, [capturedImageBlob, location, question.id, onChange, stopCamera, value]);

  useEffect(() => {
    return () => {
      stopCamera();
    };
  }, [stopCamera]);

  const handleFileChange = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      if (readOnly) return;
      const files = e.target.files;
      if (files && files.length > 0) {
        // Check file size before upload
        const maxSize = 10 * 1024 * 1024; // 10MB
        for (let i = 0; i < files.length; i++) {
          const file = files[i];
          if (file.size > maxSize) {
            const errorMsg = `File ${file.name} size (${(file.size / 1024 / 1024).toFixed(2)}MB) exceeds maximum limit of 10MB`;
            setError(errorMsg);
            e.target.value = "";
            if (inputRef.current) inputRef.current.value = "";
            return;
          }

          const allowed = question.allowedFileTypes;
          if (allowed && allowed.length > 0) {
            const isValid = allowed.some((type) => {
              const validator = FILE_VALIDATORS[type];
              return validator ? validator(file) : true;
            });
            if (!isValid) {
              const allowedLabels = allowed
                .map((type) => FILE_TYPE_LABELS[type] ?? type)
                .join(", ");
              window.alert(`Please upload a ${allowedLabels} file. (${file.name} is invalid)`);
              e.target.value = "";
              if (inputRef.current) inputRef.current.value = "";
              return;
            }

          }
        }

        try {
          setUploading(true);
          setError(null);
          setUploadProgress(null);


          const uploadedUrls: string[] = [];
          for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const result = await apiClient.uploadFile(
              file,
              "form",
              question.id,
              (progress) => {
                setUploadProgress({
                  ...progress,
                  percentage: ((i * 100) + progress.percentage) / files.length
                });
              }
            );
            const uploadedUrl = apiClient.resolveUploadedFileUrl(result);
            if (!uploadedUrl) {
              throw new Error(`File upload did not return a valid URL for ${file.name}`);
            }
            uploadedUrls.push(uploadedUrl);
          }


          let existingFiles: any[] = [];
          if (Array.isArray(value)) {
            existingFiles = value;
          } else if (value) {
            try {
              const parsed = JSON.parse(value);
              if (Array.isArray(parsed)) {
                existingFiles = parsed;
              } else {
                existingFiles = [value];
              }
            } catch {
              existingFiles = [value];
            }
          }

          onChange([...existingFiles, ...uploadedUrls]);
        } catch (err) {
          const errorMsg = err instanceof Error ? err.message : "Failed to upload file";
          setError(errorMsg);
          console.error("File upload error:", err);
          e.target.value = "";
          if (inputRef.current) {
            inputRef.current.value = "";
          }
        } finally {
          setUploading(false);
          setUploadProgress(null);
        }
      }
    },
    [onChange, question.allowedFileTypes, readOnly, value, question.id]
  );

  const handleRemoveSingleFile = useCallback((indexToRemove: number) => {
    if (readOnly) return;
    let existingFiles: any[] = [];
    if (Array.isArray(value)) {
      existingFiles = [...value];
    } else if (value) {
      try {
        const parsed = JSON.parse(value);
        if (Array.isArray(parsed)) existingFiles = [...parsed];
        else existingFiles = [value];
      } catch {
        existingFiles = [value];
      }
    }
    existingFiles.splice(indexToRemove, 1);

    if (inputRef.current) inputRef.current.value = "";

    if (existingFiles.length === 0) {
      onChange("");
    } else {
      onChange(existingFiles);
    }
    onChange("");
  }, [value, onChange, readOnly]);

  return (
    <div className="space-y-4">
      {!uploadMode ? (
        <>
          <div className="flex gap-2">
            <button
              onClick={() => setUploadMode("file")}
              disabled={readOnly}
              className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-lg bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 hover:bg-blue-200 dark:hover:bg-blue-900/50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors border border-blue-300 dark:border-blue-700"
            >
              <Upload className="w-4 h-4" />
              <span className="text-sm font-medium">Manual Upload</span>
            </button>
            {question.allowedFileTypes?.includes("image") && (
              <button
                onClick={() => setUploadMode("camera")}
                disabled={readOnly}
                className="flex-1 flex items-center justify-center gap-2 px-4 py-3 rounded-lg bg-green-100 dark:bg-green-900/30 text-green-700 dark:text-green-300 hover:bg-green-200 dark:hover:bg-green-900/50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors border border-green-300 dark:border-green-700"
              >
                <Camera className="w-4 h-4" />
                <span className="text-sm font-medium">Camera Capture</span>
              </button>
            )}
          </div>
        </>
      ) : uploadMode === "file" ? (
        <>
          <button
            type="button"
            onClick={() => setUploadMode(null)}
            className="text-sm text-blue-600 dark:text-blue-400 hover:underline mb-2"
          >
            ← Back
          </button>
          <label
            className={`flex flex-col items-center px-4 py-6 border-2 border-dashed rounded-lg ${readOnly || uploading
              ? "cursor-not-allowed bg-gray-100 dark:bg-gray-800"
              : "cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-700"
              } border-gray-300 dark:border-gray-600 ${error ? "border-red-300 dark:border-red-600" : ""}`}
          >
            {uploading ? (
              <>
                <Loader2 className="w-8 h-8 text-blue-400 mb-2 animate-spin" />
                <div className="text-center space-y-2">
                  <span className="text-sm text-blue-600 dark:text-blue-400">
                    Uploading file...
                  </span>
                  {uploadProgress && (
                    <div className="space-y-2">
                      <div className="w-full bg-gray-200 dark:bg-gray-700 rounded-full h-2">
                        <div
                          className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                          style={{ width: `${uploadProgress.percentage}%` }}
                        ></div>
                      </div>
                      <div className="text-xs text-gray-600 dark:text-gray-400 space-y-1">
                        <div className="flex justify-between">
                          <span>{uploadProgress.percentage}%</span>
                          <span>
                            {uploadProgress.timeRemaining
                              ? `${Math.floor(uploadProgress.timeRemaining / 60)}:${(uploadProgress.timeRemaining % 60).toString().padStart(2, '0')} remaining`
                              : 'Calculating...'
                            }
                          </span>
                        </div>
                        <div className="text-center">
                          {(uploadProgress.loaded / 1024 / 1024).toFixed(1)}MB / {(uploadProgress.total / 1024 / 1024).toFixed(1)}MB
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              </>
            ) : (
              <>
                <Upload className="w-8 h-8 text-gray-400 mb-2" />
                <span className="text-sm text-gray-500 dark:text-gray-500">
                  {readOnly
                    ? "File upload disabled"
                    : "Click to upload or drag and drop"}
                </span>
              </>
            )}
            <input
              ref={inputRef}
              type="file"
              accept={accept}
              onChange={handleFileChange}
              disabled={readOnly || uploading}
              className="hidden"
              required={question.required && (!value || value.length === 0)}
              multiple
            />
          </label>
        </>
      ) : uploadMode === "camera" ? (
        <>
          <button
            type="button"
            onClick={() => {
              stopCamera();
              setUploadMode(null);
              setLocation(null);
              setLocationError(null);
              setVideoReady(false);
            }}
            className="text-sm text-green-600 dark:text-green-400 hover:underline mb-2"
          >
            ← Back
          </button>

          {!cameraStream ? (
            <button
              type="button"
              onClick={startCamera}
              disabled={uploading || readOnly}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 rounded-lg bg-green-600 hover:bg-green-700 text-white disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-medium"
            >
              <Camera className="w-4 h-4" />
              Start Camera
            </button>
          ) : (
            <>
              <div className="space-y-4">
                <div className="relative bg-black rounded-lg overflow-hidden">
                  <video
                    ref={videoRef}
                    autoPlay
                    playsInline
                    muted
                    onError={(e) => {
                      console.error("Video error:", e);
                      setError("Failed to load video stream");
                    }}
                    onLoadedMetadata={() => {
                      console.log("Video metadata loaded");
                      setVideoReady(true);
                      // Force play if not already playing
                      if (videoRef.current && videoRef.current.paused) {
                        videoRef.current.play().catch(e => {
                          console.error("Auto-play failed:", e);
                        });
                      }
                    }}
                    onPlay={() => {
                      console.log("Video playing");
                      setVideoReady(true);
                    }}
                    className="w-full h-auto max-h-96"
                    style={{ display: 'block' }}
                  />
                  <canvas ref={canvasRef} className="hidden" />
                  {cameraStream && !videoReady && (
                    <div className="absolute inset-0 flex items-center justify-center bg-black">
                      <Loader2 className="w-8 h-8 text-white animate-spin" />
                      <span className="ml-2 text-white">Initializing camera...</span>
                    </div>
                  )}
                </div>

                {location ? (
                  <div className="flex items-start gap-3 p-3 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg">
                    <MapPin className="w-5 h-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-blue-900 dark:text-blue-200">
                        Location Captured
                      </p>
                      <p className="text-xs text-blue-700 dark:text-blue-300 mt-1">
                        <strong>Latitude:</strong> {location.lat.toFixed(6)}
                      </p>
                      <p className="text-xs text-blue-700 dark:text-blue-300">
                        <strong>Longitude:</strong> {location.lng.toFixed(6)}
                      </p>
                      <p className="text-xs text-blue-700 dark:text-blue-300">
                        <strong>Accuracy:</strong> ±{location.accuracy.toFixed(1)}m
                      </p>
                    </div>
                  </div>
                ) : locationLoading ? (
                  <div className="flex items-center gap-2 p-3 bg-yellow-50 dark:bg-yellow-900/20 border border-yellow-200 dark:border-yellow-700 rounded-lg">
                    <Loader2 className="w-5 h-5 text-yellow-600 dark:text-yellow-400 animate-spin flex-shrink-0" />
                    <p className="text-sm text-yellow-700 dark:text-yellow-300">
                      Getting your location...
                    </p>
                  </div>
                ) : locationError ? (
                  <div className="flex items-start gap-3 p-3 bg-yellow-50 dark:bg-yellow-900/20 border border-yellow-200 dark:border-yellow-700 rounded-lg">
                    <AlertCircle className="w-5 h-5 text-yellow-600 dark:text-yellow-400 flex-shrink-0 mt-0.5" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-yellow-900 dark:text-yellow-200">
                        {locationError}
                      </p>
                      <button
                        onClick={() => getLocation()}
                        className="text-xs text-yellow-700 dark:text-yellow-300 hover:underline mt-1 font-medium"
                      >
                        Retry Location
                      </button>
                    </div>
                  </div>
                ) : null}

                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      stopCamera();
                      setUploadMode(null);
                      setLocation(null);
                      setLocationError(null);
                      setVideoReady(false);
                    }}
                    disabled={uploading}
                    className="flex-1 px-4 py-2 rounded-lg bg-gray-200 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-300 dark:hover:bg-gray-600 disabled:opacity-50 transition-colors font-medium"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={capturePhoto}
                    disabled={uploading || !location || !videoReady}
                    className="flex-1 px-4 py-2 rounded-lg bg-green-600 hover:bg-green-700 text-white disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-medium flex items-center justify-center gap-2"
                  >
                    {showCapturePreview ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        Processing...
                      </>
                    ) : (
                      <>
                        <Camera className="w-4 h-4" />
                        Capture
                      </>
                    )}
                  </button>
                </div>
              </div>
            </>
          )}
        </>
      ) : null}

      {error && (
        <div className="flex items-center gap-2 p-3 bg-red-50 dark:bg-red-900 border border-red-200 dark:border-red-700 rounded-lg">
          <AlertCircle className="w-5 h-5 text-red-600 dark:text-red-400 flex-shrink-0" />
          <p className="text-sm text-red-600 dark:text-red-400">{error}</p>
        </div>
      )}
      {value ? (() => {
        let filesToRender: any[] = [];
        if (Array.isArray(value)) {
          filesToRender = value;
        } else if (value) {
          try {
            const parsed = JSON.parse(value);
            if (Array.isArray(parsed)) filesToRender = parsed;
            else filesToRender = [value];
          } catch {
            filesToRender = [value];
          }
        }
        if (filesToRender.length === 0) return null;

        return (
          <div className="space-y-4">
            {filesToRender.map((fileVal, index) => {
              const fileData = parseFileValue(fileVal);
              const imageUrl = fileData.type === "camera" ? fileData.url : fileVal;
              const fileName = resolveFileName(fileVal);
              const showImagePreview = isImageUrl(fileVal);
              return (
                <div key={index} className="space-y-4 animate-in fade-in slide-in-from-top-2 duration-300">
                  <div className={`flex items-center gap-3 p-4 rounded-xl border transition-all duration-300 ${isApplied
                    ? 'bg-emerald-50 dark:bg-emerald-900/10 border-emerald-500/50 ring-4 ring-emerald-500/5'
                    : 'bg-gray-50 dark:bg-gray-800 border-gray-200 dark:border-gray-700'
                    }`}>
                    <div className={`p-2 rounded-lg ${isApplied ? 'bg-emerald-500 text-white' : 'bg-blue-500 text-white'}`}>
                      {showImagePreview ? (
                        <Eye className="w-5 h-5" />
                      ) : (
                        <CheckCircle2 className="w-5 h-5" />
                      )}
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className={`text-sm font-bold truncate ${isApplied ? 'text-emerald-900 dark:text-emerald-100' : 'text-gray-900 dark:text-gray-100'}`}>
                        {fileName || "Uploaded file"}
                      </p>
                      <p className={`text-xs font-medium ${isApplied ? 'text-emerald-600 dark:text-emerald-400' : 'text-gray-500 dark:text-gray-400'}`}>
                        {fileData.type === "camera" ? "Camera captured" : showImagePreview ? "Image uploaded" : "File uploaded successfully"}
                        {isApplied && " (Auto-filled)"}
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <a
                        href={imageUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className={`inline-flex items-center gap-1.5 px-3 py-2 text-xs font-black uppercase tracking-wider rounded-lg border transition-all ${isApplied
                          ? 'text-emerald-600 bg-emerald-50 border-emerald-200 hover:bg-emerald-100 dark:text-emerald-300 dark:bg-emerald-900/40 dark:border-emerald-700'
                          : 'text-blue-600 bg-blue-50 border-blue-200 hover:bg-blue-100 dark:text-blue-300 dark:bg-blue-900/40 dark:border-blue-700'
                          }`}
                      >
                        <Eye className="w-4 h-4" />
                        View
                      </a>
                      {!readOnly && (
                        <button
                          type="button"
                          onClick={() => handleRemoveSingleFile(index)}
                          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-black uppercase tracking-wider text-red-600 bg-red-50 border border-red-200 rounded-lg hover:bg-red-100 dark:text-red-300 dark:bg-red-900/40 dark:border-red-700 transition-all"
                        >
                          <X className="w-4 h-4" />
                          Remove
                        </button>
                      )}
                    </div>

                  </div>


                  {fileData.type === "camera" && fileData.location && (
                    <div className={`flex items-start gap-3 p-3 rounded-lg border transition-all duration-300 ${isApplied
                      ? 'bg-emerald-50 dark:bg-emerald-900/20 border-emerald-200 dark:border-emerald-700'
                      : 'bg-blue-50 dark:bg-blue-900/20 border-blue-200 dark:border-blue-700'
                      }`}>
                      <MapPin className={`w-5 h-5 flex-shrink-0 mt-0.5 ${isApplied ? 'text-emerald-600 dark:text-emerald-400' : 'text-blue-600 dark:text-blue-400'}`} />
                      <div className="flex-1 min-w-0">
                        <p className={`text-sm font-bold ${isApplied ? 'text-emerald-900 dark:text-emerald-200' : 'text-blue-900 dark:text-blue-200'}`}>
                          Location Metadata
                        </p>
                        <div className={`mt-1 space-y-0.5 ${isApplied ? 'text-emerald-700 dark:text-emerald-300' : 'text-blue-700 dark:text-blue-300'}`}>
                          <p className="text-xs">
                            <strong>Latitude:</strong> {fileData.location.latitude.toFixed(6)}
                          </p>
                          <p className="text-xs">
                            <strong>Longitude:</strong> {fileData.location.longitude.toFixed(6)}
                          </p>
                          <p className="text-xs">
                            <strong>Accuracy:</strong> ±{fileData.location.accuracy.toFixed(1)}m
                          </p>
                          {fileData.timestamp && (
                            <p className="text-xs mt-1 italic">
                              <strong>Captured:</strong> {new Date(fileData.timestamp).toLocaleString()}
                            </p>
                          )}
                        </div>
                      </div>
                    </div>
                  )}
                  {showImagePreview ? (
                    <img
                      src={imageUrl}
                      alt={fileName || "Uploaded file"}
                      className="max-w-full h-auto rounded-lg max-h-64 border border-gray-200 dark:border-gray-700"
                    />
                  ) : null}
                </div>
              );
            })}
          </div>
        );
      })() : null}

      {showCapturePreview && capturedImageUrl && (
        <div className="fixed inset-0 bg-black/50 dark:bg-black/70 flex items-center justify-center z-50 p-4">
          <div className="bg-white dark:bg-gray-800 rounded-xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-auto">
            <div className="p-6 space-y-4">
              <div className="flex items-center justify-between mb-4">
                <h3 className="text-xl font-bold text-gray-900 dark:text-white">Review Captured Image</h3>
                <button
                  type="button"
                  onClick={() => {
                    setShowCapturePreview(false);
                    setCapturedImageUrl(null);
                  }}
                  className="text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200"
                >
                  ×
                </button>
              </div>

              <div className="flex flex-col items-center gap-4">
                <img
                  src={capturedImageUrl}
                  alt="Captured preview"
                  className="max-w-full h-auto rounded-lg max-h-96 border border-gray-200 dark:border-gray-700"
                />

                {location && (
                  <div className="w-full flex items-start gap-3 p-3 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg">
                    <MapPin className="w-5 h-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-blue-900 dark:text-blue-200">
                        Location Captured
                      </p>
                      <p className="text-xs text-blue-700 dark:text-blue-300 mt-1">
                        <strong>Latitude:</strong> {location.lat.toFixed(6)}
                      </p>
                      <p className="text-xs text-blue-700 dark:text-blue-300">
                        <strong>Longitude:</strong> {location.lng.toFixed(6)}
                      </p>
                      <p className="text-xs text-blue-700 dark:text-blue-300">
                        <strong>Accuracy:</strong> ±{location.accuracy.toFixed(1)}m
                      </p>
                    </div>
                  </div>
                )}

                {uploading && uploadProgress && (
                  <div className="w-full space-y-2">
                    <div className="w-full bg-gray-200 dark:bg-gray-700 rounded-full h-2">
                      <div
                        className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                        style={{ width: `${uploadProgress.percentage}%` }}
                      ></div>
                    </div>
                    <div className="text-xs text-gray-600 dark:text-gray-400 text-center">
                      Uploading... {uploadProgress.percentage}%
                    </div>
                  </div>
                )}
              </div>

              <div className="flex gap-3 justify-end pt-4 border-t border-gray-200 dark:border-gray-700">
                <button
                  type="button"
                  onClick={() => {
                    setShowCapturePreview(false);
                    setCapturedImageUrl(null);
                    setVideoReady(true);
                  }}
                  disabled={uploading}
                  className="px-4 py-2 rounded-lg bg-gray-200 dark:bg-gray-700 text-gray-700 dark:text-gray-300 hover:bg-gray-300 dark:hover:bg-gray-600 disabled:opacity-50 transition-colors font-medium"
                >
                  Retake
                </button>
                <button
                  type="button"
                  onClick={uploadCapturedImage}
                  disabled={uploading}
                  className="px-4 py-2 rounded-lg bg-green-600 hover:bg-green-700 text-white disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-medium flex items-center gap-2"
                >
                  {uploading ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      Uploading...
                    </>
                  ) : (
                    <>
                      <CheckCircle2 className="w-4 h-4" />
                      OK
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
