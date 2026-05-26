import React, { Component, ReactNode } from "react";

// Error boundary component to catch and display errors gracefully
interface ErrorBoundaryProps {
  fallback: React.ComponentType<{ error: Error; resetError: () => void }>;
  children: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("Uncaught error:", error, errorInfo);
    // You could send error to reporting service here
    // e.g., reportErrorToService(error, errorInfo);
  }

  resetError = () => {
    this.setState({ hasError: false, error: null });
  };

  render() {
    if (this.state.hasError && this.state.error) {
      const FallbackComponent = this.props.fallback;
      return <FallbackComponent error={this.state.error} resetError={this.resetError} />;
    }

    return this.props.children;
  }
}

export default ErrorBoundary;

// Simple fallback component for demonstration
export const ErrorFallback: React.FC<{ error: Error; resetError: () => void }> = ({
  error,
  resetError,
}) => {
  return (
    <div style={{ textAlign: "center", padding: "2rem" }}>
      <h2>Something went wrong.</h2>
      <p>Details: {error.message}</p>
      <button onClick={resetError} style={{ marginTop: "1rem" }}>
        Try Again
      </button>
    </div>
  );
};