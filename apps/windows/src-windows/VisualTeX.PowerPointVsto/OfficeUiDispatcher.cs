using System.Windows.Forms;

namespace VisualTeX.PowerPointVsto;

internal sealed class OfficeUiDispatcher : IDisposable
{
    private readonly Control _control;
    private readonly HashSet<System.Windows.Forms.Timer> _delayedTimers = new();

    public OfficeUiDispatcher()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("The PowerPoint add-in must initialize on the Office STA thread.");
        _control = new Control();
        _control.CreateControl();
    }

    public void Post(Action operation)
    {
        if (_control.IsDisposed || _control.Disposing) return;
        try { _control.BeginInvoke(operation); } catch (InvalidOperationException) { }
    }

    public void PostDelayed(Action operation, int delayMilliseconds)
    {
        if (_control.IsDisposed || _control.Disposing) return;

        void Schedule()
        {
            if (_control.IsDisposed || _control.Disposing) return;
            var timer = new System.Windows.Forms.Timer
            {
                Interval = Math.Max(1, delayMilliseconds),
            };
            EventHandler? onTick = null;
            onTick = (_, _) =>
            {
                timer.Stop();
                if (onTick is not null) timer.Tick -= onTick;
                _delayedTimers.Remove(timer);
                timer.Dispose();
                if (_control.IsDisposed || _control.Disposing) return;
                operation();
            };
            timer.Tick += onTick;
            _delayedTimers.Add(timer);
            timer.Start();
        }

        try
        {
            if (_control.InvokeRequired) _control.BeginInvoke(new Action(Schedule));
            else Schedule();
        }
        catch (InvalidOperationException) { }
    }

    public Task<T> InvokeAsync<T>(Func<T> operation)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        void Execute()
        {
            try { completion.TrySetResult(operation()); }
            catch (Exception error) { completion.TrySetException(error); }
        }
        if (_control.InvokeRequired) _control.BeginInvoke(new Action(Execute));
        else Execute();
        return completion.Task;
    }

    public void Dispose()
    {
        foreach (var timer in _delayedTimers.ToArray())
        {
            try { timer.Stop(); } catch { }
            timer.Dispose();
        }
        _delayedTimers.Clear();
        _control.Dispose();
    }
}
