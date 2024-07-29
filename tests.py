import win32com.client

def main():
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        presentation = app.Presentations.Open("C:\\Users\\Sabin\\OneDrive - Universitatea Babeş-Bolyai\\Metoda Cubului.pptx")
        print("Presentation loaded successfully.")
        presentation.SlideShowSettings.Run()
        input("Press Enter to exit...")
        presentation.SlideShowWindow.View.Exit()
        print("Presentation ended.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
