function [V,D,Bzeros] = core(N,M,nsteps,bens,e,L,r_ic,r_oc)
% SLOSH Code
% Following pp 60-64 of "The New Dynamic Behavior of Liquids in Moving 
% Containers" (F.T. Dodge, SWI, San Antonio, Texas), addapted from NASA
% CR-230 (D.O. Lomen). 

%% Parameters
r_cm = bens(1);
r_cM = bens(2);
z_cm = bens(3);

%% Natural frequencies (eigenvalues & eigenfunctions)

% Calculate zeros of function dJ1(r)/dr = 0
Bzeros = dBesselzero(N);

% Matrix Inicialization
A      = zeros(N,N);
B      = zeros(N,N);

% Matrix definition
h      = waitbar(0,'Please wait...');
for m = 1:N
    waitbar(m/N,h,'Please wait...')
    for n = m:N        
        A(m,n) = simp2D(@(r,z)Aint(r,z,m,n,M,Bzeros,L,r_ic,r_oc), r_cm, r_cM, z_cm, L, nsteps, nsteps);
        B(m,n) = Bint(m,n,M,Bzeros,e);
    end
end
close(h)

% Symmetry
for m = 1:N
    for n = 1:N
        A(n,m) = A(m,n);
        B(n,m) = B(m,n);
    end
end

% EIG
[V,D] = eig(A,B);
