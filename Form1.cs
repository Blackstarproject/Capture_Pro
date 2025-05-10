// Created by: Justin Linwood Ross | 5/9/2025 |GitHub: https://github.com/Rythorian77?tab=repositories
// For Other Tutorials Go To: https://www.youtube.com/@justinlinwoodrossakarythor63/videos
//
// This application, 'Capture_Pro', is designed for real-time motion detection using a webcam.
// It captures video frames, applies image processing filters to detect motion, and can trigger alerts,
// save snapshots, and log events when motion is detected.
//
// The application is well-engineered to handle memory and resource management effectively.
// The consistent use of 'using' statements for disposable objects created within method scopes,
// coupled with explicit disposal of long-lived class members and proper shutdown procedures for external resources like the camera,
// significantly minimizes the risk of memory leaks.
// This provides a solid foundation for a stable application from a resource management perspective.

using AForge.Imaging;             // For image processing filters like Grayscale, Difference, Threshold, BlobCounter.
using AForge.Imaging.Filters;     // Specific namespace for AForge image filters.
using AForge.Video;               // Base namespace for AForge video capture.
using AForge.Video.DirectShow;    // For capturing video from DirectShow compatible devices (webcams).
using System;                     // Provides fundamental classes and base types.
using System.Collections.Generic; // For List<T>.
using System.Drawing;             // For graphics objects like Bitmap, Graphics, Pen, Rectangle, Color.
using System.Drawing.Imaging;     // For ImageFormat (used for saving images).
using System.IO;                  // For file and directory operations (Path, Directory, File).
using System.Media;               // For playing system sounds.
using System.Windows.Forms;       // For Windows Forms UI elements (Form, PictureBox, Button, ComboBox, Label, MessageBox).
using System.Diagnostics;         // Added for System.Diagnostics.EventLog and Debug.

namespace Capture_Pro // This is the application's primary namespace. Consider renaming if 'Application\'s Name' is different.
{
    /// <summary>
    /// The main form for the Capture_Pro motion detection application.
    /// Handles camera initialization, frame processing, motion detection logic,
    /// and UI updates, ensuring robust resource management.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Private Members - Variables

        #region Camera Related
        /// <summary>
        /// Collection of available video input devices (webcams) on the system.
        /// </summary>
        private FilterInfoCollection videoDevices;
        /// <summary>
        /// Represents the video capture device (webcam) currently in use.
        /// </summary>
        private VideoCaptureDevice videoSource;
        #endregion

        #region Motion Detection Related
        /// <summary>
        /// Used to count and extract information about detected blobs (connected components)
        /// in the motion map, representing areas of motion.
        /// </summary>
        private BlobCounter blobCounter;
        /// <summary>
        /// Stores the grayscale version of the previously processed frame.
        /// Used by the Difference filter to detect changes between consecutive frames.
        /// This object is disposed and re-assigned with each new frame.
        /// </summary>
        private Bitmap previousFrame;
        #endregion

        #region Filters
        /// <summary>
        /// A pre-initialized grayscale filter using the BT709 algorithm.
        /// This is a static instance as the filter itself is stateless and reusable.
        /// </summary>
        private readonly Grayscale grayscaleFilter = Grayscale.CommonAlgorithms.BT709;
        /// <summary>
        /// Filter used to calculate the absolute difference between the current frame
        /// and the previous frame, highlighting areas of change (motion).
        /// This object is re-initialized or its OverlayImage updated as previousFrame changes.
        /// </summary>
        private Difference differenceFilter;
        /// <summary>
        /// Filter used to convert the difference map into a binary image,
        /// where pixels above a certain threshold are white (motion) and others are black (no motion).
        /// The threshold value is configurable.
        /// </summary>
        private Threshold thresholdFilter;
        #endregion

        #region Drawing Pens
        /// <summary>
        /// Pen used for drawing rectangles around detected motion blobs on the video feed.
        /// Disposed explicitly on application shutdown or camera stop.
        /// </summary>
        private Pen greenPen = new Pen(Color.Green, 2);
        #endregion

        #region Automatic Enhancements & Settings
        /// <summary>
        /// Flag indicating whether active motion is currently being detected.
        /// </summary>
        private bool _isMotionActive = false;
        /// <summary>
        /// Timestamp of the last time an alert sound was played.
        /// Used to implement a cooldown period for alerts.
        /// </summary>
        private DateTime _lastAlertTime = DateTime.MinValue;
        /// <summary>
        /// Timestamp of the last time a motion snapshot was saved.
        /// Used to implement a cooldown period for saving images.
        /// </summary>
        private DateTime _lastSaveTime = DateTime.MinValue;
        /// <summary>
        /// Timestamp when the current continuous motion detection event started.
        /// Used for logging motion event durations.
        /// </summary>
        private DateTime _currentMotionStartTime = DateTime.MinValue;
        /// <summary>
        /// Timestamp of the very last frame where motion was detected.
        /// Used to determine when motion has ceased after a period of no detection.
        /// </summary>
        private DateTime _lastMotionDetectionTime = DateTime.MinValue;

        // Fixed Settings (No UI for adjustment in this version, can be made configurable later)
        /// <summary>
        /// The threshold value for the Threshold filter. Pixels with intensity difference
        /// above this value are considered motion. Lower values increase sensitivity.
        /// </summary>
        private readonly int _motionThreshold = 15;

        // Fixed Blob size filtering thresholds
        /// <summary>
        /// Minimum width (in pixels) a detected blob must have to be considered significant motion.
        /// </summary>
        private readonly int _minBlobWidth = 20;
        /// <summary>
        /// Minimum height (in pixels) a detected blob must have to be considered significant motion.
        /// </summary>
        private readonly int _minBlobHeight = 20;
        /// <summary>
        /// Maximum width (in pixels) a detected blob can have. Larger blobs might be noise or light changes.
        /// </summary>
        private readonly int _maxBlobWidth = 500;
        /// <summary>
        /// Maximum height (in pixels) a detected blob can have. Larger blobs might be noise or light changes.
        /// </summary>
        private readonly int _maxBlobHeight = 500;

        /// <summary>
        /// Cooldown duration after an alert sound is played before another can be played.
        /// Prevents continuous beeping during prolonged motion.
        /// </summary>
        private readonly TimeSpan _alertCooldown = TimeSpan.FromSeconds(5);
        /// <summary>
        /// Cooldown duration after an image is saved before another can be saved.
        /// Prevents an excessive number of images during prolonged motion.
        /// </summary>
        private readonly TimeSpan _saveCooldown = TimeSpan.FromSeconds(3);

        /// <summary>
        /// The directory path where detected motion images will be saved.
        /// Defaults to a 'Detected Images' folder on the user's Desktop.
        /// </summary>
        private readonly string _saveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Detected Images");
        /// <summary>
        /// The file path for the application's motion event log.
        /// Located in the application's startup directory.
        /// </summary>
        private readonly string _logFilePath = Path.Combine(Application.StartupPath, "motion_log.txt");

        /// <summary>
        /// Flag to enable or disable saving motion event snapshots.
        /// </summary>
        private readonly bool _saveEventsEnabled = true;
        /// <summary>
        /// Flag to enable or disable playing an alert sound on motion detection.
        /// </summary>
        private readonly bool _alertSoundEnabled = true;
        /// <summary>
        /// Flag to enable or disable updating the status label on the UI for motion events.
        /// </summary>
        private readonly bool _alertLabelEnabled = true;
        /// <summary>
        /// Flag to enable or disable logging motion events to a file and EventLog.
        /// </summary>
        private readonly bool _logEventsEnabled = true;

        // --- Added for robust logging ---
        /// <summary>
        /// Counts consecutive failures when attempting to write to the motion log file.
        /// </summary>
        private int _consecutiveLogErrors = 0;
        /// <summary>
        /// The maximum number of consecutive log file write failures before considering
        /// file logging critically failed and potentially falling back to EventLog exclusively.
        /// </summary>
        private const int _maxConsecutiveLogErrors = 5;
        /// <summary>
        /// Flag to indicate if file logging has encountered a critical failure (e.g., permissions issues)
        /// beyond which it should not attempt to write to the file again.
        /// </summary>
        private bool _isFileLoggingCriticallyFailed = false;
        /// <summary>
        /// Timestamp of the last time a critical logging error message box was shown.
        /// Used to rate-limit message box pop-ups to avoid spamming the user.
        /// </summary>
        private DateTime _lastLogErrorMessageBoxTime = DateTime.MinValue;
        /// <summary>
        /// Cooldown duration for displaying a critical logging error message box.
        /// </summary>
        private readonly TimeSpan _logErrorMessageBoxCooldown = TimeSpan.FromMinutes(5);
        // --- End Added for robust logging ---

        #endregion

        #region Region of Interest (ROI) Variables
        /// <summary>
        /// Represents the selected Region of Interest (ROI) as a rectangle.
        /// If null, the entire frame is processed for motion.
        /// Initialized to cover the entire videoPictureBox by default.
        /// </summary>
        private Rectangle? _roiSelection = null;
        /// <summary>
        /// Pen used for drawing the Region of Interest rectangle on the video feed.
        /// Disposed explicitly on application shutdown or camera stop.
        /// </summary>
        private Pen _roiPen = new Pen(Color.Red, 2);
        #endregion

        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the Form1 class.
        /// Sets up UI components, initializes motion detection objects,
        /// and loads available camera devices.
        /// </summary>
        public Form1()
        {
            InitializeComponent(); // Initializes components defined in the designer file (e.g., buttons, picture boxes).

            // Set the _roiSelection to match the initial size of the videoPictureBox.
            // This makes the entire video feed the default ROI for motion detection.
            _roiSelection = new Rectangle(0, 0, videoPictureBox.Width, videoPictureBox.Height);

            // Subscribe to the FormClosing event to ensure resources are properly released
            // when the application is closed by the user.
            FormClosing += new FormClosingEventHandler(Form1_FormClosing);

            // Initialize the BlobCounter, which is used to find and analyze motion blobs.
            blobCounter = new BlobCounter();
            // Initialize the Threshold filter with the default motion threshold.
            thresholdFilter = new Threshold(_motionThreshold);

            // Initially disable the Start and Stop buttons until cameras are loaded.
            startButton.Enabled = false;
            stopButton.Enabled = false;

            LoadCameraDevices(); // Attempt to discover and list available camera devices.
        }
        #endregion

        #region Camera Management
        /// <summary>
        /// Discovers and populates the camera combo box with available video input devices.
        /// Handles cases where no cameras are found or an error occurs during discovery.
        /// </summary>
        private void LoadCameraDevices()
        {
            try
            {
                // Update status label on the UI, ensuring it's thread-safe via Invoke.
                if (statusLabel != null && statusLabel.InvokeRequired)
                {
                    statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "Loading cameras..."; });
                }
                else if (statusLabel != null)
                {
                    statusLabel.Text = "Loading cameras...";
                }

                // Create a collection of all video input devices found on the system.
                videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

                if (videoDevices.Count == 0)
                {
                    // If no cameras are found, inform the user and disable the start button.
                    MessageBox.Show("No video sources found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    startButton.Enabled = false;
                    // Update status label.
                    if (statusLabel != null && statusLabel.InvokeRequired)
                    {
                        statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "No cameras found."; });
                    }
                    else if (statusLabel != null)
                    {
                        statusLabel.Text = "No cameras found.";
                    }
                    LogMotionEvent("WARNING: No video sources found."); // Log this significant event.
                }
                else
                {
                    // If cameras are found, clear any existing items and add each device's name to the combo box.
                    cameraComboBox.Items.Clear();
                    foreach (FilterInfo device in videoDevices)
                    {
                        cameraComboBox.Items.Add(device.Name);
                    }

                    // Select the first camera by default.
                    cameraComboBox.SelectedIndex = 0;
                    // Enable the Start button as a camera is available.
                    startButton.Enabled = true;

                    // Update status label with the number of cameras found.
                    if (statusLabel != null && statusLabel.InvokeRequired)
                    {
                        statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = $"Found {videoDevices.Count} camera(s). Ready."; });
                    }
                    else if (statusLabel != null)
                    {
                        statusLabel.Text = $"Found {videoDevices.Count} camera(s). Ready.";
                    }
                    LogMotionEvent($"INFO: Found {videoDevices.Count} camera(s). Ready."); // Log successful camera detection.
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions during camera loading, display an error message,
                // and disable the start button to prevent further issues.
                MessageBox.Show("Failed to load video sources: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                startButton.Enabled = false;
                // Update status label.
                if (statusLabel != null && statusLabel.InvokeRequired)
                {
                    statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "Error loading cameras."; });
                }
                else if (statusLabel != null)
                {
                    statusLabel.Text = "Error loading cameras.";
                }
                LogMotionEvent($"ERROR: Failed to load video sources: {ex.Message}"); // Log the error for debugging.
            }
        }

        /// <summary>
        /// Stops the currently running video source and disposes of all associated resources
        /// to prevent memory leaks and ensure a clean shutdown. This method is designed
        /// to be robust against individual disposal failures.
        /// </summary>
        private void StopCamera()
        {
            // Attempt to stop the video source first.
            try
            {
                if (videoSource != null && videoSource.IsRunning)
                {
                    // Unsubscribe from the NewFrame event to prevent further processing
                    // after the camera has stopped.
                    videoSource.NewFrame -= new NewFrameEventHandler(VideoSource_NewFrame);
                    // Signal the video source to stop capturing frames.
                    videoSource.SignalToStop();
                    // Wait for the video source to completely stop. This is crucial for proper shutdown.
                    videoSource.WaitForStop();
                    // Nullify the videoSource reference to allow garbage collection.
                    videoSource = null;
                    LogMotionEvent("INFO: Video source signaled to stop and waited for stop.");
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Exception while stopping video source: {ex.Message}");
                // Attempt to continue with resource disposal even if stopping the video source
                // failed partially, as other resources might still need to be released.
            }

            // Centralized disposal logic with individual try-catch blocks for robustness.
            // This ensures that if one resource fails to dispose, others can still be released.
            try
            {
                if (videoPictureBox.Image != null)
                {
                    // Dispose the image currently displayed in the PictureBox to release its memory.
                    videoPictureBox.Image.Dispose();
                    videoPictureBox.Image = null; // Clear the image reference.
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose videoPictureBox.Image: {ex.Message}");
            }

            try
            {
                if (previousFrame != null)
                {
                    // Dispose the previous frame bitmap used for motion detection.
                    previousFrame.Dispose();
                    previousFrame = null; // Clear the reference.
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose previousFrame: {ex.Message}");
            }

            // AForge filter objects like Difference and Threshold typically don't hold
            // unmanaged resources that require explicit Dispose(). They are usually
            // re-initialized or their properties updated. Nullifying them here
            // helps with garbage collection and ensures a clean state reset for restart.
            try
            {
                if (differenceFilter != null)
                {
                    differenceFilter = null;
                }
            }
            catch (Exception ex) // Catching just in case, though unlikely for nulling references
            {
                LogMotionEvent($"ERROR: Failed to reset differenceFilter: {ex.Message}");
            }

            try
            {
                if (thresholdFilter != null)
                {
                    thresholdFilter = null;
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to reset thresholdFilter: {ex.Message}");
            }

            try
            {
                if (blobCounter != null)
                {
                    // BlobCounter also typically doesn't require explicit dispose,
                    // but nulling helps for state reset.
                    blobCounter = null;
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to reset blobCounter: {ex.Message}");
            }

            // Dispose of GDI+ objects (Pens). These *do* hold unmanaged resources and must be disposed.
            try
            {
                if (greenPen != null)
                {
                    greenPen.Dispose();
                    greenPen = null; // Clear the reference.
                }
                // Re-initialize the pen for subsequent starts if needed, or ensure it's created in the constructor.
                // For consistency, if it's a class member, it's better to recreate it than to leave it null.
                greenPen = new Pen(Color.Green, 2); // Re-create for next use
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose or re-initialize greenPen: {ex.Message}");
            }

            try
            {
                if (_roiPen != null)
                {
                    _roiPen.Dispose();
                    _roiPen = null; // Clear the reference.
                }
                _roiPen = new Pen(Color.Red, 2); // Re-create for next use
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose or re-initialize _roiPen: {ex.Message}");
            }

            // Reset motion detection state flags.
            _isMotionActive = false;

            // Update UI status label.
            if (statusLabel != null && statusLabel.InvokeRequired)
            {
                statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "Camera stopped. Ready."; });
            }
            else if (statusLabel != null)
            {
                statusLabel.Text = "Camera stopped. Ready.";
            }

            // Update button and combo box states for restarting.
            startButton.Enabled = true;
            stopButton.Enabled = false;
            cameraComboBox.Enabled = true;

            LogMotionEvent("INFO: Camera stopped. Resources disposed and reset.");
        }
        #endregion

        #region Event Handlers
        /// <summary>
        /// Handles the click event for the 'Start' button.
        /// Initializes and starts the selected video capture device.
        /// Resets motion detection state variables.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event data.</param>
        private void StartButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate that a camera is selected and available.
                if (videoDevices == null || videoDevices.Count == 0 || cameraComboBox.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select a camera.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (statusLabel != null) statusLabel.Text = "Select a camera.";
                    LogMotionEvent("WARNING: Start button clicked with no camera selected or found.");
                    return; // Exit the method if no camera is selected.
                }

                // Re-initialize threshold filter if it's null or its value has changed
                // (though _motionThreshold is a readonly field, this check is good practice if it were configurable).
                if (thresholdFilter == null || thresholdFilter.ThresholdValue != _motionThreshold)
                {
                    thresholdFilter = new Threshold(_motionThreshold);
                }

                // Create a new VideoCaptureDevice instance using the moniker string of the selected device.
                videoSource = new VideoCaptureDevice(videoDevices[cameraComboBox.SelectedIndex].MonikerString);
                // Unsubscribe defensively to prevent multiple subscriptions if Start is clicked repeatedly
                // without a full application restart or proper StopCamera call.
                videoSource.NewFrame -= new NewFrameEventHandler(VideoSource_NewFrame);
                // Subscribe to the NewFrame event, which is triggered whenever a new video frame is available.
                videoSource.NewFrame += new NewFrameEventHandler(VideoSource_NewFrame);
                // Begin video capture.
                videoSource.Start();

                // Dispose of any previous frame stored from a prior run and clear the reference.
                if (previousFrame != null)
                {
                    previousFrame.Dispose();
                    previousFrame = null;
                }
                // Reset the difference filter as the previous frame has been reset.
                differenceFilter = null;

                // Reset all motion detection state variables to their initial values.
                _lastAlertTime = DateTime.MinValue;
                _lastSaveTime = DateTime.MinValue;
                _isMotionActive = false;
                _currentMotionStartTime = DateTime.MinValue;
                _lastMotionDetectionTime = DateTime.MinValue;

                // Update UI element states.
                startButton.Enabled = false;
                stopButton.Enabled = true;
                cameraComboBox.Enabled = false;

                // Update status label.
                if (_alertLabelEnabled && statusLabel != null) statusLabel.Text = "Camera started. Detecting motion...";
                LogMotionEvent("INFO: Application started. Camera activated.");
            }
            catch (Exception ex)
            {
                // Handle any errors during camera startup.
                MessageBox.Show("Error starting camera: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Re-enable start button and disable stop button on error.
                startButton.Enabled = true;
                stopButton.Enabled = false;
                cameraComboBox.Enabled = true;
                // Attempt to signal the video source to stop if it somehow started partially.
                if (videoSource != null && videoSource.IsRunning)
                {
                    videoSource.SignalToStop();
                }
                videoSource = null; // Ensure videoSource reference is cleared.
                if (_alertLabelEnabled && statusLabel != null) statusLabel.Text = "Error starting camera.";
                LogMotionEvent($"ERROR: Failed to start camera: {ex.Message}"); // Log the error.
            }
        }

        /// <summary>
        /// Event handler for when a new frame is received from the video source.
        /// This method is crucial for real-time processing and motion detection.
        /// It clones the frame, processes it for motion, updates the UI, and handles
        /// motion-triggered events (alerts, saving snapshots).
        /// </summary>
        /// <param name="sender">The video source that sent the frame.</param>
        /// <param name="eventArgs">Arguments containing the new frame.</param>
        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            // 'currentFrame' will hold a clone of the original frame for display.
            Bitmap currentFrame = null;
            // 'frameToProcess' will hold a clone specifically for motion analysis,
            // allowing drawing on 'currentFrame' for display without affecting analysis.
            Bitmap frameToProcess = null;

            try
            {
                // Clone the incoming frame to work with it.
                // The original eventArgs.Frame is managed by AForge and should not be disposed here.
                currentFrame = (Bitmap)eventArgs.Frame.Clone();
                // Clone again for the motion processing pipeline.
                frameToProcess = (Bitmap)currentFrame.Clone();

                // Process the 'frameToProcess' for motion detection.
                // This method will draw rectangles on 'frameToProcess' if motion is found.
                bool motionDetected = ProcessFrameForMotion(frameToProcess);

                // Update the UI (videoPictureBox and statusLabel) on the UI thread.
                // This is critical because NewFrame events occur on a separate thread.
                if (videoPictureBox.InvokeRequired)
                {
                    videoPictureBox.Invoke((MethodInvoker)delegate
                    {
                        try
                        {
                            // Dispose of the previously displayed image in the PictureBox
                            // to prevent GDI+ memory leaks.
                            videoPictureBox.Image?.Dispose();
                            // Assign the processed frame (which might have motion rectangles) to the PictureBox.
                            videoPictureBox.Image = frameToProcess;
                            // Handle motion-related events (alerts, saving) using the original frame for saving.
                            HandleMotionEvents(motionDetected, currentFrame); // Pass the original clone for saving.
                        }
                        catch (Exception invokeEx)
                        {
                            // Log errors occurring within the Invoke delegate.
                            System.Diagnostics.Debug.WriteLine($"Error in VideoPictureBox Invoke delegate: {invokeEx.Message}");
                            LogMotionEvent($"ERROR: Exception in VideoPictureBox Invoke delegate: {invokeEx.Message}");
                            // Ensure disposal of bitmaps if an error occurs during UI update.
                            currentFrame?.Dispose();
                            frameToProcess?.Dispose();
                            // Attempt to stop the camera on a critical error that prevents UI updates.
                            StopCamera();
                        }
                    });
                }
                else // If not invoked (should generally not happen for NewFrame, but as fallback).
                {
                    videoPictureBox.Image?.Dispose();
                    videoPictureBox.Image = frameToProcess;
                    HandleMotionEvents(motionDetected, currentFrame);
                }
            }
            catch (Exception ex)
            {
                // Log errors occurring in the main NewFrame handler logic.
                System.Diagnostics.Debug.WriteLine($"Error in NewFrame (main thread logic): {ex.Message}");
                LogMotionEvent($"ERROR: Exception in NewFrame handler: {ex.Message}");
                // Ensure all created bitmaps are disposed in case of an error.
                currentFrame?.Dispose();
                currentFrame = null;
                frameToProcess?.Dispose();
                frameToProcess = null;
                // Attempt to stop the camera on critical error that prevents further frame processing.
                if (videoPictureBox.InvokeRequired)
                {
                    videoPictureBox.Invoke((MethodInvoker)delegate
                    {
                        StopCamera();
                    });
                }
                else
                {
                    StopCamera();
                }
            }
        }

        /// <summary>
        /// Handles the click event for the 'Stop' button.
        /// Initiates the camera stopping procedure and logs the event.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event data.</param>
        private void StopButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Call the robust StopCamera method to halt video and release resources.
                StopCamera();
                LogMotionEvent("INFO: Application stopped. Camera deactivated.");
                // Close the form, which will also trigger the Form1_FormClosing event.
                this.Close();
            }
            catch (Exception ex)
            {
                // Handle any errors during the stopping process, inform the user,
                // and re-enable appropriate UI elements.
                MessageBox.Show("Error stopping camera: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogMotionEvent($"ERROR: Failed to stop camera: {ex.Message}");
                startButton.Enabled = true;
                stopButton.Enabled = false;
                cameraComboBox.Enabled = true;
                // If stopping failed critically, force application exit to prevent hanging resources.
                Application.Exit();
            }
        }

        /// <summary>
        /// Event handler for the form closing event.
        /// Ensures that camera resources are properly released before the application exits.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event data, indicating how the form is closing.</param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Call StopCamera to ensure all resources are released when the form closes.
            // StopCamera already has robust internal error handling for disposal.
            // Any critical failures in StopCamera will be logged internally.
            StopCamera();
            LogMotionEvent("INFO: Application is closing.");
        }
        #endregion

        #region Motion Processing Logic
        /// <summary>
        /// Processes a single video frame to detect motion.
        /// Applies grayscale, difference, and threshold filters, then uses BlobCounter
        /// to identify motion regions. Draws green rectangles around detected blobs.
        /// Handles Region of Interest (ROI) if defined.
        /// </summary>
        /// <param name="frame">The current frame (Bitmap) to be processed.
        /// This bitmap will be modified to draw motion rectangles.</param>
        /// <returns>True if motion is detected, false otherwise.</returns>
        private bool ProcessFrameForMotion(Bitmap frame)
        {
            bool motionDetected = false;

            try
            {
                // Apply grayscale filter to the current frame.
                // 'using' ensures the grayscale image is disposed after its scope.
                using (Bitmap grayCurrentFrame = grayscaleFilter.Apply(frame))
                {
                    // Proceed with motion detection only if a previous frame exists for comparison.
                    if (previousFrame != null)
                    {
                        // Initialize Difference filter if it's null, or update its OverlayImage.
                        // The OverlayImage is the 'background' or 'previous' frame for comparison.
                        if (differenceFilter == null)
                        {
                            differenceFilter = new Difference(previousFrame);
                        }
                        else
                        {
                            differenceFilter.OverlayImage = previousFrame;
                        }

                        // Apply the difference filter to get a 'motion map' highlighting changes.
                        using (Bitmap motionMap = differenceFilter.Apply(grayCurrentFrame))
                        {
                            // Initialize Threshold filter if null or update its value.
                            if (thresholdFilter == null)
                            {
                                thresholdFilter = new Threshold(_motionThreshold);
                            }
                            else if (thresholdFilter.ThresholdValue != _motionThreshold)
                            {
                                thresholdFilter.ThresholdValue = _motionThreshold; // Update threshold if it changed
                            }

                            // Apply the threshold filter to convert the motion map into a binary image.
                            // 'using' ensures the binaryMotionMap is disposed.
                            using (Bitmap binaryMotionMap = thresholdFilter.Apply(motionMap))
                            {
                                Bitmap imageToProcessForBlobs = binaryMotionMap; // Default to full binary motion map.
                                Rectangle currentRoi = Rectangle.Empty; // Initialize ROI for blob translation.

                                // If an ROI is defined, crop the binary motion map to the ROI.
                                if (_roiSelection.HasValue)
                                {
                                    currentRoi = _roiSelection.Value;

                                    // Ensure ROI is within image bounds before attempting to crop.
                                    if (currentRoi.X >= 0 && currentRoi.Y >= 0 &&
                                        currentRoi.Right <= binaryMotionMap.Width &&
                                        currentRoi.Bottom <= binaryMotionMap.Height)
                                    {
                                        // Create a Crop filter using the defined ROI.
                                        // AForge.NET's Crop filter is a struct, so it doesn't need Dispose().
                                        Crop cropFilter = new Crop(currentRoi);
                                        // Apply the crop filter and ensure the resulting bitmap is disposed after use.
                                        imageToProcessForBlobs = cropFilter.Apply(binaryMotionMap);
                                    }
                                    else
                                    {
                                        // Log a warning if ROI is out of bounds and fallback to processing the full frame.
                                        System.Diagnostics.Debug.WriteLine("Warning: ROI is out of image bounds. Processing full frame.");
                                        LogMotionEvent("WARNING: ROI is out of image bounds. Processing full frame.");
                                        imageToProcessForBlobs = binaryMotionMap; // Fallback to full frame.
                                        currentRoi = Rectangle.Empty; // Clear ROI to indicate full frame processing (no offset needed for blobs).
                                    }
                                }

                                // Process the (possibly cropped) binary motion map with the BlobCounter.
                                blobCounter.ProcessImage(imageToProcessForBlobs);

                                // Get the rectangles of all detected blobs.
                                Rectangle[] allBlobsRects = blobCounter.GetObjectsRectangles();

                                List<Rectangle> filteredBlobs = new List<Rectangle>();
                                // Iterate through detected blobs to filter by size (min/max width/height).
                                foreach (Rectangle blobRect in allBlobsRects)
                                {
                                    Rectangle translatedBlobRect = blobRect;
                                    // If an ROI was applied, translate the blob coordinates back to the original frame's coordinate system.
                                    if (currentRoi != Rectangle.Empty)
                                    {
                                        translatedBlobRect.X += currentRoi.X;
                                        translatedBlobRect.Y += currentRoi.Y;
                                    }

                                    // Apply size filtering to the blob.
                                    if (translatedBlobRect.Width >= _minBlobWidth && translatedBlobRect.Height >= _minBlobHeight &&
                                        translatedBlobRect.Width <= _maxBlobWidth && translatedBlobRect.Height <= _maxBlobHeight)
                                    {
                                        filteredBlobs.Add(translatedBlobRect);
                                    }
                                }

                                // Dispose cropped image if it was created (i.e., not equal to binaryMotionMap).
                                // This is important if a new bitmap was created by the Crop filter.
                                if (imageToProcessForBlobs != binaryMotionMap && imageToProcessForBlobs != null)
                                {
                                    imageToProcessForBlobs.Dispose();
                                    imageToProcessForBlobs = null;
                                }

                                // If any blobs passed the size filter, motion is detected.
                                if (filteredBlobs.Count > 0)
                                {
                                    motionDetected = true;

                                    // Draw green rectangles around the detected motion blobs on the original color frame.
                                    // 'using' ensures the Graphics object is disposed.
                                    using (Graphics g = Graphics.FromImage(frame))
                                    {
                                        foreach (Rectangle blobRect in filteredBlobs)
                                        {
                                            g.DrawRectangle(greenPen, blobRect);
                                        }
                                    }
                                }
                            } // binaryMotionMap is disposed here
                        } // motionMap is disposed here
                        previousFrame.Dispose(); // Dispose the old previous frame as it's about to be replaced.
                    }
                    // Set the current grayscale frame as the new previous frame for the next iteration.
                    // Clone it to prevent issues if grayCurrentFrame is disposed.
                    previousFrame = (Bitmap)grayCurrentFrame.Clone();
                } // grayCurrentFrame is disposed here

                // Draw ROI on the final frame (the original color frame passed in) if a region is defined.
                // This overlay helps the user visualize the active detection area.
                if (_roiSelection.HasValue)
                {
                    using (Graphics g = Graphics.FromImage(frame))
                    {
                        g.DrawRectangle(_roiPen, _roiSelection.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Exception during motion processing (ProcessFrameForMotion): {ex.Message}");
                // Re-throw the exception so the NewFrame handler can catch it and potentially stop the camera,
                // preventing continuous errors.
                throw;
            }

            return motionDetected;
        }
        #endregion

        #region Motion Event Handling
        /// <summary>
        /// Handles actions to be taken when motion is detected or ceases.
        /// Triggers alerts, saves snapshots, and updates status based on cooldowns and settings.
        /// </summary>
        /// <param name="motionDetected">True if motion was detected in the current frame, false otherwise.</param>
        /// <param name="originalFrame">The original, unprocessed color frame, used for saving snapshots.</param>
        private void HandleMotionEvents(bool motionDetected, Bitmap originalFrame)
        {
            try
            {
                if (motionDetected)
                {
                    // If motion was not previously active, mark the start of a new motion event.
                    if (!_isMotionActive)
                    {
                        _currentMotionStartTime = DateTime.Now;
                        LogMotionEvent($"MOTION STARTED at {_currentMotionStartTime:yyyy-MM-dd HH:mm:ss.fff}");
                        if (_alertLabelEnabled && statusLabel != null)
                        {
                            statusLabel.Text = "MOTION DETECTED!";
                        }
                    }
                    _isMotionActive = true; // Set motion active flag.
                    _lastMotionDetectionTime = DateTime.Now; // Update last detection time.

                    // Play alert sound if enabled and the alert cooldown has passed.
                    if (_alertSoundEnabled && DateTime.Now - _lastAlertTime > _alertCooldown)
                    {
                        SystemSounds.Asterisk.Play();
                        _lastAlertTime = DateTime.Now;
                        LogMotionEvent("INFO: Alert sound played.");
                    }

                    // Save motion snapshot if enabled and the save cooldown has passed.
                    if (_saveEventsEnabled && DateTime.Now - _lastSaveTime > _saveCooldown)
                    {
                        try
                        {
                            // Create the save directory if it doesn't exist.
                            if (!Directory.Exists(_saveDirectory))
                            {
                                Directory.CreateDirectory(_saveDirectory);
                                LogMotionEvent($"INFO: Created save directory: {_saveDirectory}");
                            }

                            // Generate a unique filename with timestamp.
                            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                            string filename = $"motion_{timestamp}.jpg";
                            string fullPath = Path.Combine(_saveDirectory, filename);

                            // Clone the original frame before saving to avoid GDI+ issues if the
                            // original frame is simultaneously being used elsewhere (e.g., PictureBox).
                            using (Bitmap frameToSave = (Bitmap)originalFrame.Clone())
                            {
                                frameToSave.Save(fullPath, ImageFormat.Jpeg);
                            }

                            _lastSaveTime = DateTime.Now; // Update last save time.

                            LogMotionEvent($"SAVED snapshot: {filename}");

                            if (_alertLabelEnabled && statusLabel != null)
                            {
                                statusLabel.Text = $"Saved motion event: {filename}";
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error saving motion event: {ex.Message}");
                            LogMotionEvent($"ERROR: Failed to save snapshot: {ex.Message}");
                            if (_alertLabelEnabled && statusLabel != null)
                            {
                                statusLabel.Text = "Error saving motion event!";
                            }
                        }
                    }
                }
                else // No motion detected in this frame
                {
                    // If motion was previously active but hasn't been detected for a short grace period,
                    // consider motion as stopped. This prevents rapid toggling of "motion detected" status.
                    if (_isMotionActive && DateTime.Now - _lastMotionDetectionTime > TimeSpan.FromMilliseconds(500)) // A small grace period (e.g., 0.5 seconds)
                    {
                        _isMotionActive = false; // Mark motion as inactive.
                        LogMotionEvent($"MOTION STOPPED at {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}. Duration: {(DateTime.Now - _currentMotionStartTime).TotalSeconds:F1}s");
                        if (_alertLabelEnabled && statusLabel != null)
                        {
                            statusLabel.Text = "No motion detected.";
                        }
                    }
                    else if (!_isMotionActive) // If motion was never active, just keep displaying no motion.
                    {
                        if (_alertLabelEnabled && statusLabel != null)
                        {
                            statusLabel.Text = "No motion detected.";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Exception in HandleMotionEvents: {ex.Message}");
            }
            finally
            {
                // Ensure the original frame bitmap passed into this handler is disposed
                // to release its memory, as it was cloned specifically for this processing cycle.
                originalFrame?.Dispose();
            }
        }
        #endregion

        #region Utility Methods
        /// <summary>
        /// Logs motion-related events to a file and, in case of critical file logging failures,
        /// falls back to the Windows Event Log. Implements error handling and rate-limiting
        /// for error message boxes to prevent spamming the user.
        /// </summary>
        /// <param name="message">The log message.</param>
        private void LogMotionEvent(string message)
        {
            // If logging is globally disabled or critically failed for file logging, do not attempt file logging.
            if (!_logEventsEnabled || _isFileLoggingCriticallyFailed)
            {
                // If file logging failed critically, but logging is still enabled,
                // we still want to log to the Event Log as a fallback.
                // This path handles cases where _isFileLoggingCriticallyFailed is true
                // and we are still attempting to log.
                if (_logEventsEnabled && _isFileLoggingCriticallyFailed)
                {
                    try
                    {
                        string eventSource = "CaptureProMotionDetection";
                        string eventLogName = "Application"; // Default Application log

                        // Check if the event source exists. Creating it requires administrative privileges
                        // the first time on a machine. For production, consider creating this during installation.
                        if (!EventLog.SourceExists(eventSource))
                        {
                            EventLog.CreateEventSource(eventSource, eventLogName);
                        }
                        // Log the message to the Windows Event Log as a Warning, indicating file logging issues.
                        EventLog.WriteEntry(eventSource,
                            $"FILE LOGGING CRITICALLY FAILED: {message}",
                            EventLogEntryType.Warning);
                    }
                    catch (Exception eventLogEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"CRITICAL ERROR: Could not write to Event Log either: {eventLogEx.Message}");
                        // At this point, all logging attempts have failed. No further action, as it's an emergency.
                    }
                }
                return; // Exit if logging is disabled or critically failed for file.
            }

            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: {message}{Environment.NewLine}";

            try
            {
                string logDirectory = Path.GetDirectoryName(_logFilePath);
                // Ensure the log directory exists.
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // Use a lock to ensure thread-safe file writing, especially important
                // as NewFrame events occur on a separate thread and multiple log calls
                // could happen concurrently.
                lock (this) // Using 'this' (the Form instance) as the lock object.
                {
                    File.AppendAllText(_logFilePath, logEntry);
                }

                // Reset consecutive error count on successful log write.
                _consecutiveLogErrors = 0;
                // If logging was previously critically failed, and now it's successful, reset the flag.
                if (_isFileLoggingCriticallyFailed)
                {
                    _isFileLoggingCriticallyFailed = false;
                    System.Diagnostics.Debug.WriteLine("File logging recovered from critical failure.");
                }
            }
            catch (Exception ex)
            {
                _consecutiveLogErrors++; // Increment error counter.
                System.Diagnostics.Debug.WriteLine($"Error logging event to file (Attempt {_consecutiveLogErrors}): {ex.Message}");

                // Fallback to Windows Event Log for critical errors or repeated failures.
                try
                {
                    string eventSource = "CaptureProMotionDetection";
                    string eventLogName = "Application"; // Or a custom log like "CapturePro Logs"

                    // IMPORTANT: Creating an EventLog source (the first time it's run on a machine)
                    // often requires administrative privileges. If the application is not run with
                    // admin rights, this line might throw an UnauthorizedAccessException.
                    // For production, consider creating the event source during application installation.
                    if (!EventLog.SourceExists(eventSource))
                    {
                        EventLog.CreateEventSource(eventSource, eventLogName);
                    }

                    // Log to Event Log. Use Error type if _isFileLoggingCriticallyFailed.
                    EventLog.WriteEntry(eventSource,
                        $"FILE LOGGING ERROR (Attempt {_consecutiveLogErrors}): {message} - Details: {ex.Message}",
                        _consecutiveLogErrors >= _maxConsecutiveLogErrors ? EventLogEntryType.Error : EventLogEntryType.Warning);
                }
                catch (Exception eventLogEx)
                {
                    // If even EventLog fails, output to Debug console as a last resort.
                    System.Diagnostics.Debug.WriteLine($"CRITICAL ERROR: Could not write to Windows Event Log while handling file logging error: {eventLogEx.Message}");
                }

                // Check for critical logging failure threshold.
                if (_consecutiveLogErrors >= _maxConsecutiveLogErrors)
                {
                    _isFileLoggingCriticallyFailed = true; // Mark file logging as critically failed.

                    // Rate-limit the message box to avoid spamming the user.
                    if (DateTime.Now - _lastLogErrorMessageBoxTime > _logErrorMessageBoxCooldown)
                    {
                        MessageBox.Show($"CRITICAL ERROR: File logging to '{_logFilePath}' has failed multiple times. " +
                                        "Future events will be logged to the Windows Event Log only. " +
                                        $"Please check file permissions or disk space. Last error: {ex.Message}",
                                        "Critical Logging Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        _lastLogErrorMessageBoxTime = DateTime.Now;
                    }
                }
            }
        }
        #endregion
    }
}