// WebSocket message types

// Supported Office applications
export type AppType = 'word' | 'excel' | 'powerpoint' | 'outlook';

export interface RegisterMessage {
  type: 'register';
  payload: {
    app: AppType;
    filename: string;
    path: string;
  };
}

export interface RegisteredMessage {
  type: 'registered';
  payload: {
    id: string;
  };
}

export interface ExecuteMessage {
  type: 'execute';
  id: string;
  payload: {
    code: string;
    timeout: number;
  };
}

export interface ResultMessage {
  type: 'result';
  id: string;
  payload: {
    success: boolean;
    result?: unknown;
    error?: {
      message: string;
      code?: string;
      stack?: string;
    };
  };
}

export interface ConsoleMessage {
  type: 'console';
  payload: {
    level: 'log' | 'warn' | 'error' | 'info';
    message: string;
    timestamp: string;
  };
}

export type AddInMessage = RegisterMessage | ResultMessage | ConsoleMessage;
export type BridgeMessage = RegisteredMessage | ExecuteMessage;

// Session registry types (app-agnostic naming)

export interface ConnectedSession {
  id: string;
  app: AppType;
  filename: string;
  path: string;
  connectedAt: Date;
  lastActivity: Date;
  status: 'connected' | 'disconnected';
  disconnectedAt?: Date;
}

export interface SessionInfo {
  id: string;
  app: AppType;
  filename: string;
  path: string;
  connectedAt: string;
  lastActivity: string;
  status: 'connected' | 'disconnected';
}

// Backward compatibility aliases
export type ConnectedDocument = ConnectedSession;
export type DocumentInfo = SessionInfo;

// Server info response
export interface ServerInfo {
  port: number;
  documents: number;
}
