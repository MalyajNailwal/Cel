import React, { useState } from 'react';

interface ConfirmModalProps {
  isOpen: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  confirmVariant?: 'danger' | 'warning' | 'primary';
  onConfirm: () => void;
  onCancel: () => void;
}

export const ConfirmModal: React.FC<ConfirmModalProps> = ({
  isOpen,
  title,
  message,
  confirmLabel = 'Confirm',
  cancelLabel = 'Cancel',
  confirmVariant = 'primary',
  onConfirm,
  onCancel,
}) => {
  if (!isOpen) return null;

  const buttonStyles = {
    danger: 'bg-red-500 hover:bg-red-600 text-white',
    warning: 'bg-amber-500 hover:bg-amber-600 text-white',
    primary: 'bg-[#217346] hover:bg-[#1a5c38] text-white',
  };

  return (
    <div className="fixed inset-y-0 right-0 z-50 flex items-center justify-center pointer-events-none">
      <div className="absolute inset-0 bg-black/30 pointer-events-auto" onClick={onCancel} />
      <div className="relative bg-white rounded-l-lg shadow-2xl w-80 max-h-[90vh] overflow-auto pointer-events-auto mr-4 my-auto">
        <div className="p-4">
          <h3 className="text-sm font-bold text-gray-900 mb-2">{title}</h3>
          <p className="text-xs text-gray-600 mb-4 whitespace-pre-wrap max-h-32 overflow-y-auto">{message}</p>
          <div className="flex gap-2">
            <button
              onClick={onCancel}
              className="flex-1 px-3 py-1.5 text-xs text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md font-medium transition-colors"
            >
              {cancelLabel}
            </button>
            <button
              onClick={onConfirm}
              className={`flex-1 px-3 py-1.5 text-xs rounded-md font-medium transition-colors ${buttonStyles[confirmVariant]}`}
            >
              {confirmLabel}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

interface ConfirmState {
  isOpen: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  confirmVariant?: 'danger' | 'warning' | 'primary';
  onConfirm: () => void;
  onCancel: () => void;
}

export const useConfirm = () => {
  const [confirmState, setConfirmState] = useState<ConfirmState | null>(null);

  const confirm = (
    message: string,
    options?: {
      title?: string;
      confirmLabel?: string;
      cancelLabel?: string;
      confirmVariant?: 'danger' | 'warning' | 'primary';
    }
  ): Promise<boolean> => {
    return new Promise((resolve) => {
      setConfirmState({
        isOpen: true,
        title: options?.title || 'Confirm',
        message,
        confirmLabel: options?.confirmLabel,
        cancelLabel: options?.cancelLabel,
        confirmVariant: options?.confirmVariant || 'primary',
        onConfirm: () => {
          setConfirmState(null);
          resolve(true);
        },
        onCancel: () => {
          setConfirmState(null);
          resolve(false);
        },
      });
    });
  };

  const Modal = confirmState ? (
    <ConfirmModal
      isOpen={confirmState.isOpen}
      title={confirmState.title}
      message={confirmState.message}
      confirmLabel={confirmState.confirmLabel}
      cancelLabel={confirmState.cancelLabel}
      confirmVariant={confirmState.confirmVariant}
      onConfirm={confirmState.onConfirm}
      onCancel={confirmState.onCancel}
    />
  ) : null;

  return { confirm, Modal };
};