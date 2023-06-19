function phi = phi(n,M,r,z,Bzeros,L)
% Function phi from A.3 and A.4
% n: Mode n
% M: Number of shallow tank functions
% r,z: Cylindrical coordinates
% Bzeros: Zeros of Besssel function of first kind
% L: Dimensionless heigh of liquid over CoM

if n <= M
    phi   = r.^(2*n-1);
else
    phi   = besselj(1,Bzeros(n-M)*r).*exp(-Bzeros(n-M)*(L-z));
end