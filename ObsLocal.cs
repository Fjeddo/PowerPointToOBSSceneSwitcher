using System;
using System.Collections.Generic;
using System.Linq;
using OBS.WebSocket.NET;

namespace PowerPointToOBSSceneSwitcher
{
    public class ObsLocal : IDisposable
	{
		private bool _disposedValue;

		private ObsWebSocket _obs;
		private List<string> _validScenes;

		private string _defaultScene;

        public ObsLocal()
        {
			_obs = new ObsWebSocket();
		}

        public void Connect(string password = "")
		{
			_obs.Connect("ws://127.0.0.1:4444", password);
		}

		public string DefaultScene
        {
            get => _defaultScene;
            set
			{
				if (_validScenes.Contains(value))
				{
					_defaultScene = value;
				}
                else
                {
                    Console.WriteLine($"Scene named {value} does not exist and cannot be set as default");
                }
			}
        }

		public bool ChangeScene(string scene)
        {
			if (!_validScenes.Contains(scene))
			{
                Console.WriteLine($"Scene named {scene} does not exist");

				if (string.IsNullOrEmpty(_defaultScene))
				{
                    Console.WriteLine("No default scene has been set!");

					return false;
				}
			
				scene = _defaultScene;
			}

			_obs.Api.SetCurrentScene(scene);

			return true;
        }

		public void GetScenes()
        {
			var allScene = _obs.Api.GetSceneList();
			var list = allScene.Scenes.Select(s => s.Name).ToList();
            Console.WriteLine("┌───────────────────────────────────────");
			Console.WriteLine("|  Valid Scenes:");
			foreach(var l in list)
            {
                Console.WriteLine($"|  {l}");
            }
            Console.WriteLine("└───────────────────────────────────────");
			_validScenes = list;
        }

		public bool StartRecording(bool quietly = false)
		{
            try
            {
                _obs.Api.StartRecording();
            }
            catch (Exception exception)
            {
                if (!quietly)
                {
                    Console.WriteLine($"Start recording failed: {exception.Message}");
                }
                return false;
            }

            return true;
		}

		public bool StopRecording()
		{
            try
            {
                _obs.Api.StopRecording();
                Console.WriteLine("Stopped recording");
            }
            catch (Exception exception)
			{
				Console.WriteLine($"Stop recording failed: {exception.Message}");
                return false;
			}

            return true;
		}

        public bool PauseRecording(bool quietly = false)
        {
            try
            {
                _obs.Api.PauseRecording();
            }
            catch (Exception exception)
            {
                if (!quietly)
                {
                    Console.WriteLine($"Pause recording failed: {exception.Message}");
                }
                return false;
            }

            return true;
        }

        public bool ResumeRecording(bool quietly = false)
        {
            try
            {
                _obs.Api.ResumeRecording();
            }
            catch (Exception exception)
            {
                if (!quietly)
                {
                    Console.WriteLine($"Resume recording failed: {exception.Message}");
                }
                return false;
            }

            return true;
        }

        public bool StartPauseResumeRecording(bool quietly = false)
        {
            if (StartRecording(quietly))
            {
                Console.WriteLine("Started recording");
                return true;
            }

            if (PauseRecording(quietly))
            {
                Console.WriteLine("Paused recording");
                return true;
            }

            if (ResumeRecording(quietly))
            {
                Console.WriteLine("Resumed recording");
                return true;
            }

            Console.Error("Tried to start, pause and resume recording, all failed");
            return false;
        }

        protected virtual void Dispose(bool disposing)
		{
			if (!_disposedValue)
			{
				if (disposing)
				{
					// TODO: dispose managed state (managed objects)
				}

				_obs.Disconnect();
				_obs = null;

				_disposedValue = true;
			}
		}

		~ObsLocal()
		{
			// Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
			Dispose(false);
		}

		public void Dispose()
		{
			// Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
			Dispose(true);
			GC.SuppressFinalize(this);
		}
	}
}