export type CarryForwardResetContext = {
  currentPage: number;
  currentSection: number;
  paragraphElement: Element;
  paragraphText: string;
  previousNumberingInfo: string;
  styleId?: string;
};

export type CarryForwardPolicy = {
  enabled: boolean;
  shouldReset: (context: CarryForwardResetContext) => boolean;
};

export type NumberingBehaviorOptions = {
  carryForward?: {
    enabled?: boolean;
    shouldReset?: (context: CarryForwardResetContext) => boolean;
  };
};

export const defaultCarryForwardPolicy: CarryForwardPolicy = {
  enabled: true,
  shouldReset: () => false,
};

export const resolveCarryForwardPolicy = (
  options?: NumberingBehaviorOptions
): CarryForwardPolicy => ({
  enabled: options?.carryForward?.enabled ?? defaultCarryForwardPolicy.enabled,
  shouldReset: options?.carryForward?.shouldReset ?? defaultCarryForwardPolicy.shouldReset,
});
