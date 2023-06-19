function [K,w,Vc,Mm,Ln,Hm,eta0n_xn,M0] = derivedparams(rho,flvol,H,V,D,g,L,a,B,hk,Ph_a,intb,selm)
% Pre-calculation
N  = length(hk);
D  = diag(D);

% Nondimensional frequency
K  = D.*(L/a);

% Natural frequencies
w = sqrt((D.*g)./a);

% Characteristic velocity
Vc = w(1)*a;

% Length of pendulum
Ln = a./D;

% Fluid mass
M  = flvol*rho;

% gamma parameter
gamma = zeros(N,1);

for n = 1:N
  gamma(n) = pi*a*a^2/flvol * sum(sum(V(:,n)*V(:,n)'.*B));
end

% b parameter
b = (pi*a^3./(flvol*gamma)).*V'*B(1,:)';

% Mass of slosh modes
Mm = M.*D.*gamma.*b.^2;

% h parameter
h  = (2*pi*a.^3)./(flvol.*gamma.*D).*V'*hk;

% Height of the slosh mass pendulum hinge above the bottom of the tank
Hm =-H+(a).*(1./D + h./(b));

% Ratio of slosh wave amplitude to slosh mass amplitude
eta0n_xn = (g/a).*Mm.*D;

% Inert mass of fluid
Mtot = sum(Mm(1:selm));
M0 = M-Mtot;

%Inert mass height
Hcg = abs(intb(3));
H0 = (M*Hcg - sum(Mm(1:selm).*(Hm(1:selm)-Ln(1:selm))))/M0;

% Height of inert fluid over cg

%Add M0 and H0 to beginning of Data
Mm = [M0; Mm];
Hm = [H0; Hm];
Ln = [0; Ln];
eta0n_xn = [0; eta0n_xn];

% Inert mass of fluid
% M0 = M*(1-sum(gamma(1:N).*b(1:N).^2.*K(1:N))); %%% !!! Select nº of modes

% Height of inert fluid over cg

